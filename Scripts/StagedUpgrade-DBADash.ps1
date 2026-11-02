#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Central', 'Remote')]
    [string]$DeploymentRole,

    [string]$ManifestUrl = 'https://addsys.com/DBADash/StagedUpgrade-Manifest.json',
    [string]$InstallPath = 'C:\ADD_Sys\DBADash',
    [string]$ServiceName = 'DBADashService',
    [string]$LogPath = 'C:\ADD_Sys\DBADashUpgrade\DBADashUpgrade.log',
    [string]$ValidationReceiptPath = 'C:\ADD_Sys\DBADashUpgrade\CentralValidation.json',
    [string]$RemoteInstanceIdentifier,

    # Required for Central. Use Integrated Security under the scheduled-task identity.
    [string]$RepositoryConnectionString,

    [ValidateRange(1, 86400)][int]$SqlCommandTimeoutSeconds = 30,
    [ValidateRange(1, 100)][int]$SqlRetryCount = 12,
    [ValidateRange(0, 3600)][int]$SqlRetryDelaySeconds = 10,
    [ValidateRange(1, 86400)][int]$ServiceTimeoutSeconds = 300,
    [ValidateRange(0, 86400)][int]$StabilizationSeconds = 60,

    # Optional environment-specific scalar query. It must return 1 to pass.
    [string]$RepositoryCustomHealthQuery = '',

    [switch]$AllowDowngrade
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# $GitHubRepository = 'ADD-Systems/dba-dash'
$GitHubRepository = 'trimble-oss/dba-dash'
$ServiceExeName = 'DBADashService.exe'
$PreserveFiles = @('ServiceConfig.json', 'PerformanceCountersCustom.xml')
$TempRoot = $null
$UpgradeStarted = $false
$BackupCreated = $false
$PreviousVersion = $null

#region Utility functions
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR')][string]$Level = 'INFO'
    )

    $directory = Split-Path -Parent $LogPath
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }

    $line = '{0} [{1}] [{2}] {3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $DeploymentRole.ToUpperInvariant(), $Message
    $line | Tee-Object -FilePath $LogPath -Append
}

function ConvertTo-DBADashVersion {
    param([Parameter(Mandatory)][object]$Value)

    $text = ([string]$Value).Trim().TrimStart('v')
    if ($text -notmatch '^\d+\.\d+\.\d+(?:\.\d+)?$') {
        throw "Invalid DBA Dash version '$text'."
    }

    return [version]$text
}

function Get-InstalledDBADashVersion {
    param([Parameter(Mandatory)][string]$ExecutablePath)

    if (-not (Test-Path -LiteralPath $ExecutablePath)) {
        throw "Executable not found: $ExecutablePath"
    }

    $info = (Get-Item -LiteralPath $ExecutablePath).VersionInfo
    foreach ($candidate in @($info.ProductVersion, $info.FileVersion)) {
        if ($candidate -and $candidate -match '(\d+\.\d+\.\d+(?:\.\d+)?)') {
            return [version]$Matches[1]
        }
    }

    throw "Cannot determine version from $ExecutablePath."
}

function Restore-InstallFromBackup {
    param(
        [Parameter(Mandatory)][string]$BackupPath,
        [Parameter(Mandatory)][string]$InstallPath
    )

    if (-not (Test-Path -LiteralPath $BackupPath)) {
        return
    }

    if (-not (Test-Path -LiteralPath $InstallPath)) {
        New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
    }

    Get-ChildItem -LiteralPath $InstallPath -Force | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
    }

    Copy-Item -Path (Join-Path $BackupPath '*') -Destination $InstallPath -Recurse -Force
}

function Wait-ServiceStatus {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet('Running', 'Stopped')][string]$Status,
        [int]$TimeoutSeconds = 300
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $svc = Get-Service -Name $Name -ErrorAction Stop
        if ($svc.Status.ToString() -eq $Status) {
            return
        }
        Start-Sleep -Seconds 2
    } while ((Get-Date) -lt $deadline)

    throw "Service '$Name' did not reach '$Status' within $TimeoutSeconds seconds."
}

function Wait-ServiceStability {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateRange(0, 86400)][int]$Seconds
    )

    Write-Log "Beginning $Seconds-second service stabilization check."
    $deadline = (Get-Date).AddSeconds($Seconds)
    do {
        if ((Get-Service -Name $Name -ErrorAction Stop).Status -ne 'Running') {
            throw "Service '$Name' stopped during stabilization."
        }

        if ($Seconds -gt 0) {
            $remainingSeconds = ($deadline - (Get-Date)).TotalSeconds
            if ($remainingSeconds -gt 0) {
                Start-Sleep -Seconds ([Math]::Min(5, [Math]::Max(1, [Math]::Ceiling($remainingSeconds))))
            }
        }
    } while ((Get-Date) -lt $deadline)

    Write-Log 'Service stabilization check passed.'
}

function Test-ServiceExists {
    param([Parameter(Mandatory)][string]$Name)

    return $null -ne (Get-Service -Name $Name -ErrorAction SilentlyContinue)
}

function Invoke-SqlScalar {
    param(
        [Parameter(Mandatory)][string]$ConnectionString,
        [Parameter(Mandatory)][string]$CommandText,
        [int]$TimeoutSeconds = 30
    )

    $connection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $CommandText
        $command.CommandTimeout = $TimeoutSeconds
        return $command.ExecuteScalar()
    }
    finally {
        if ($connection) {
            $connection.Dispose()
        }
    }
}

function Test-SqlRepositoryHealth {
    param(
        [Parameter(Mandatory)][string]$ConnectionString,
        [int]$RetryCount = 12,
        [int]$RetryDelaySeconds = 10,
        [int]$CommandTimeoutSeconds = 30,
        [string]$CustomHealthQuery = ''
    )

    if (-not [string]::IsNullOrWhiteSpace($CustomHealthQuery) -and -not (Test-RepositoryCustomHealthQuery -Query $CustomHealthQuery)) {
        throw 'RepositoryCustomHealthQuery contains unsupported SQL elements; only a single read-only scalar SELECT is allowed.'
    }

    $baseQuery = @'
SET NOCOUNT ON;
SELECT CASE
 WHEN DB_NAME() IS NOT NULL
  AND DATABASEPROPERTYEX(DB_NAME(), 'Status') = 'ONLINE'
  AND DATABASEPROPERTYEX(DB_NAME(), 'Updateability') = 'READ_WRITE'
  AND HAS_DBACCESS(DB_NAME()) = 1
  AND (SELECT COUNT_BIG(*) FROM sys.tables WHERE is_ms_shipped = 0) > 0
 THEN 1 ELSE 0 END;
'@

    $lastError = $null
    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            Write-Log "SQL repository health attempt $attempt of $RetryCount."
            $result = Invoke-SqlScalar -ConnectionString $ConnectionString -CommandText $baseQuery -TimeoutSeconds $CommandTimeoutSeconds
            if ([int]$result -ne 1) {
                throw 'Repository is not ONLINE, READ_WRITE, accessible, or contains no user tables.'
            }
            Write-Log 'Repository connectivity, database state, access, and user-table checks passed.'

            if (-not [string]::IsNullOrWhiteSpace($CustomHealthQuery)) {
                $customResult = Invoke-SqlScalar -ConnectionString $ConnectionString -CommandText $CustomHealthQuery -TimeoutSeconds $CommandTimeoutSeconds
                if ([int]$customResult -ne 1) {
                    throw "Custom repository health query returned '$customResult'; expected 1."
                }
                Write-Log 'Custom repository health query passed.'
            }

            return $true
        }
        catch {
            $lastError = $_.Exception.Message
            Write-Log "SQL repository health attempt failed: $lastError" 'WARNING'
            if ($attempt -lt $RetryCount -and $RetryDelaySeconds -gt 0) {
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }

    throw "SQL repository health checks failed after $RetryCount attempts. Last error: $lastError"
}

function Test-RepositoryCustomHealthQuery {
    param([Parameter(Mandatory)][string]$Query)

    $trimmed = $Query.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $false
    }

    if ($trimmed -match '(?is)(;\s*(?:--|/\*|SELECT|WAITFOR|INSERT|UPDATE|DELETE|EXEC|ALTER|DROP|CREATE|TRUNCATE|MERGE|BEGIN|COMMIT|ROLLBACK))') {
        return $false
    }

    if ($trimmed -notmatch '(?is)^\s*SELECT\b') {
        return $false
    }

    $blockedTokens = @('INSERT', 'UPDATE', 'DELETE', 'EXEC', 'ALTER', 'DROP', 'CREATE', 'TRUNCATE', 'MERGE', 'BEGIN', 'COMMIT', 'ROLLBACK')
    foreach ($token in $blockedTokens) {
        if ($trimmed -match "(?is)\b$([regex]::Escape($token))\b") {
            return $false
        }
    }

    return $true
}

function Get-VerifiedPackageHash {
    param(
        [Parameter(Mandatory)][version]$Version,
        [Parameter(Mandatory)][string]$ExpectedHash
    )

    $zipName = "DBADash_$($Version.ToString()).zip"
    $downloadUrl = "https://github.com/$GitHubRepository/releases/download/$Version/$zipName"
    $zipPath = Join-Path ([IO.Path]::GetTempPath()) ("DBADash-{0}-{1}.zip" -f $Version, [guid]::NewGuid().ToString('N'))

    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing -TimeoutSec 900
        $actualHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToUpperInvariant()
        if ($actualHash -ne $ExpectedHash) {
            throw "SHA-256 mismatch. Expected $ExpectedHash; actual $actualHash."
        }

        return $actualHash
    }
    finally {
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    }
}

function Test-CentralHealth {
    param(
        [Parameter(Mandatory)][version]$ExpectedVersion,
        [Parameter(Mandatory)][string]$ConnectionString
    )

    $exe = Join-Path $InstallPath $ServiceExeName
    $config = Join-Path $InstallPath 'ServiceConfig.json'

    Wait-ServiceStatus -Name $ServiceName -Status Running -TimeoutSeconds $ServiceTimeoutSeconds
    if ((Get-InstalledDBADashVersion -ExecutablePath $exe) -ne $ExpectedVersion) {
        throw 'Installed executable version does not match the central target version.'
    }

    if (-not (Test-Path -LiteralPath $config)) {
        throw 'ServiceConfig.json is missing.'
    }

    Get-Content -LiteralPath $config -Raw | ConvertFrom-Json | Out-Null
    Write-Log 'ServiceConfig.json exists and is valid JSON.'

    Wait-ServiceStability -Name $ServiceName -Seconds $StabilizationSeconds

    Test-SqlRepositoryHealth -ConnectionString $ConnectionString `
        -RetryCount $SqlRetryCount `
        -RetryDelaySeconds $SqlRetryDelaySeconds `
        -CommandTimeoutSeconds $SqlCommandTimeoutSeconds `
        -CustomHealthQuery $RepositoryCustomHealthQuery | Out-Null
}

function Write-ValidationReceipt {
    param(
        [Parameter(Mandatory)][version]$Version,
        [Parameter(Mandatory)][string]$PackageHash
    )

    $directory = Split-Path -Parent $ValidationReceiptPath
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }

    [ordered]@{
        version       = $Version.ToString()
        status        = 'Passed'
        validatedUtc  = (Get-Date).ToUniversalTime().ToString('o')
        computerName  = $env:COMPUTERNAME
        packageSha256 = $PackageHash
    } | ConvertTo-Json | Set-Content -LiteralPath $ValidationReceiptPath -Encoding UTF8

    Write-Log "Central validation receipt written to $ValidationReceiptPath."
}
#endregion

try {
    Write-Log 'Starting DBA Dash deployment check.'
    $manifestContent = (Invoke-WebRequest -Uri $ManifestUrl -UseBasicParsing -TimeoutSec 60).Content
    if ([string]::IsNullOrWhiteSpace($manifestContent)) {
        throw 'Deployment manifest is empty.'
    }
    $manifest = $manifestContent | ConvertFrom-Json

    if ($null -eq $manifest.PSObject.Properties['schemaVersion'] -or [int]$manifest.schemaVersion -ne 3) {
        throw 'Unsupported or missing deployment manifest schemaVersion; expected 3.'
    }

    $expiresAt = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse([string]$manifest.expiresUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$expiresAt)) {
        throw 'Deployment manifest has an invalid or missing expiresUtc value.'
    }
    if ($expiresAt -le [DateTimeOffset]::UtcNow) {
        throw "Deployment manifest expired at $($expiresAt.ToString('o'))."
    }

    if ($DeploymentRole -eq 'Central') {
        if ($null -eq $manifest.PSObject.Properties['central']) {
            throw 'Deployment manifest is missing central configuration.'
        }
        $desiredText = $manifest.central.version
    }
    else {
        if ([string]::IsNullOrWhiteSpace($RemoteInstanceIdentifier)) {
            throw 'RemoteInstanceIdentifier is required for Remote deployments.'
        }

        if ($null -eq $manifest.PSObject.Properties['remoteInstances']) {
            throw 'Deployment manifest is missing remoteInstances.'
        }

        $remoteInstances = @($manifest.remoteInstances | Where-Object { [string]$_.identifier -eq $RemoteInstanceIdentifier })
        if ($remoteInstances.Count -ne 1) {
            throw "Expected exactly one remote instance named '$RemoteInstanceIdentifier'; found $($remoteInstances.Count)."
        }

        $remoteInstance = $remoteInstances[0]
        if ($remoteInstance.enabled -ne $true) {
            throw "Remote deployment is disabled for instance '$RemoteInstanceIdentifier'."
        }

        $desiredText = $remoteInstance.version
    }
    $desiredVersion = ConvertTo-DBADashVersion $desiredText
    $desiredVersionText = $desiredVersion.ToString()
    Write-Log "Desired version: $desiredVersionText."

    if ($RepositoryCustomHealthQuery -and -not (Test-RepositoryCustomHealthQuery -Query $RepositoryCustomHealthQuery)) {
        throw 'RepositoryCustomHealthQuery is invalid. It must be a single read-only scalar SELECT statement.'
    }

    if (-not $PSCmdlet.ShouldProcess("DBA Dash at $InstallPath", "upgrade to $desiredVersionText")) {
        Write-Log 'Upgrade skipped because -WhatIf was specified or the operation was declined.'
        exit 0
    }

    # Remote hard gate: central validation must be Passed and at least the requested remote version.
    if ($DeploymentRole -eq 'Remote') {
        if ($null -eq $manifest.central.validation) {
            throw 'Remote deployment blocked: central.validation is missing.'
        }

        if ([string]$manifest.central.validation.status -ne 'Passed') {
            throw 'Remote deployment blocked: central validation status is not Passed.'
        }

        $validatedVersion = ConvertTo-DBADashVersion $manifest.central.validation.version
        $centralVersion = ConvertTo-DBADashVersion $manifest.central.version
        if ($desiredVersion -gt $centralVersion) {
            throw 'Remote deployment blocked: remoteVersion is newer than centralVersion.'
        }

        if ($desiredVersion -gt $validatedVersion) {
            throw "Remote deployment blocked: central is validated only through $validatedVersion."
        }

        Write-Log "Remote promotion gate passed; central validated version is $validatedVersion."
    }

    $packageProperty = $manifest.packages.PSObject.Properties | Where-Object Name -EQ $desiredVersionText | Select-Object -First 1
    if ($null -eq $packageProperty) {
        throw "No package metadata exists for $desiredVersionText."
    }

    $hashText = ([string]$packageProperty.Value.sha256).Trim()
    if ($hashText -match '^[A-F0-9]{64}$') {
        $expectedHash = $hashText.ToUpperInvariant()
    }
    else {
        throw "Invalid SHA-256 for $desiredVersionText."
    }

    $serviceExe = Join-Path $InstallPath $ServiceExeName
    $currentVersion = Get-InstalledDBADashVersion -ExecutablePath $serviceExe
    $PreviousVersion = $currentVersion
    Write-Log "Installed version: $currentVersion."

    if ($currentVersion -eq $desiredVersion) {
        if ($DeploymentRole -eq 'Central') {
            if ([string]::IsNullOrWhiteSpace($RepositoryConnectionString)) {
                throw 'RepositoryConnectionString is required for Central validation.'
            }

            $actualHash = Get-VerifiedPackageHash -Version $desiredVersion -ExpectedHash $expectedHash
            Test-CentralHealth -ExpectedVersion $desiredVersion -ConnectionString $RepositoryConnectionString
            Write-ValidationReceipt -Version $desiredVersion -PackageHash $actualHash
        }
        else {
            Wait-ServiceStability -Name $ServiceName -Seconds $StabilizationSeconds
        }

        Write-Log 'Already at desired version; validation completed.'
        exit 0
    }

    if (($currentVersion -gt $desiredVersion) -and -not $AllowDowngrade) {
        Write-Log "Installed version is newer than desired; downgrade blocked." 'WARNING'
        exit 0
    }

    $zipName = "DBADash_$desiredVersionText.zip"
    $downloadUrl = "https://github.com/$GitHubRepository/releases/download/$desiredVersionText/$zipName"
    $TempRoot = Join-Path $env:TEMP ("DBADashUpgrade-{0}-{1}" -f $desiredVersionText, [guid]::NewGuid().ToString('N'))
    $zipPath = Join-Path $TempRoot $zipName
    $extractPath = Join-Path $TempRoot 'Extracted'
    $backupPath = Join-Path $TempRoot 'Rollback'
    $configPath = Join-Path $TempRoot 'Config'
    New-Item -Path $extractPath, $backupPath, $configPath -ItemType Directory -Force | Out-Null

    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing -TimeoutSec 900
    $actualHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actualHash -ne $expectedHash) {
        throw "SHA-256 mismatch. Expected $expectedHash; actual $actualHash."
    }
    Write-Log 'SHA-256 verification passed.'

    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractPath -Force
    $extractedExe = Join-Path $extractPath $ServiceExeName
    if ((Get-InstalledDBADashVersion -ExecutablePath $extractedExe) -ne $desiredVersion) {
        throw 'Downloaded executable version does not match the manifest.'
    }

    $UpgradeStarted = $true

    if (Test-ServiceExists -Name $ServiceName) {
        Stop-Service -Name $ServiceName -ErrorAction Stop
        Wait-ServiceStatus -Name $ServiceName -Status Stopped -TimeoutSeconds $ServiceTimeoutSeconds
    }
    else {
        Write-Log "Service '$ServiceName' was not found; continuing without stopping it." 'WARNING'
    }

    Copy-Item -Path (Join-Path $InstallPath '*') -Destination $backupPath -Recurse -Force
    foreach ($name in $PreserveFiles) {
        $source = Join-Path $InstallPath $name
        if (Test-Path -LiteralPath $source) {
            Copy-Item -LiteralPath $source -Destination (Join-Path $configPath $name) -Force
        }
    }
    $BackupCreated = $true

    Get-ChildItem -LiteralPath $InstallPath -Force | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
    }

    Copy-Item -Path (Join-Path $extractPath '*') -Destination $InstallPath -Recurse -Force
    foreach ($name in $PreserveFiles) {
        $saved = Join-Path $configPath $name
        if (Test-Path -LiteralPath $saved) {
            Copy-Item -LiteralPath $saved -Destination (Join-Path $InstallPath $name) -Force
        }
    }

    if ((Get-InstalledDBADashVersion -ExecutablePath $serviceExe) -ne $desiredVersion) {
        throw 'Post-copy version verification failed.'
    }

    if (Test-ServiceExists -Name $ServiceName) {
        Start-Service -Name $ServiceName -ErrorAction Stop
        Wait-ServiceStatus -Name $ServiceName -Status Running -TimeoutSeconds $ServiceTimeoutSeconds
    }
    else {
        Write-Log "Service '$ServiceName' was not found after upgrade; startup was skipped." 'WARNING'
    }

    if ($DeploymentRole -eq 'Central') {
        if ([string]::IsNullOrWhiteSpace($RepositoryConnectionString)) {
            throw 'RepositoryConnectionString is required for Central validation.'
        }

        Test-CentralHealth -ExpectedVersion $desiredVersion -ConnectionString $RepositoryConnectionString
        Write-ValidationReceipt -Version $desiredVersion -PackageHash $actualHash
    }
    else {
        Wait-ServiceStability -Name $ServiceName -Seconds $StabilizationSeconds
    }

    $UpgradeStarted = $false
    Write-Log "Deployment succeeded; running DBA Dash $desiredVersion."
    try {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force -ErrorAction Stop
        $TempRoot = $null
    }
    catch {
        Write-Log "Deployment succeeded, but temporary files could not be removed: $($_.Exception.Message)" 'WARNING'
    }
    exit 0
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    if ($UpgradeStarted -and $TempRoot) {
        try {
            $backupPath = Join-Path $TempRoot 'Rollback'
            if (Test-ServiceExists -Name $ServiceName) {
                Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            }

            if ($BackupCreated) {
                Restore-InstallFromBackup -BackupPath $backupPath -InstallPath $InstallPath
                Write-Log "Rollback restored version $PreviousVersion." 'WARNING'
            }
            else {
                Write-Log 'Install backup was not completed; installation files were not changed.' 'WARNING'
            }

            if (Test-ServiceExists -Name $ServiceName) {
                Start-Service -Name $ServiceName -ErrorAction Stop
                Wait-ServiceStatus -Name $ServiceName -Status Running -TimeoutSeconds $ServiceTimeoutSeconds
            }
        }
        catch {
            Write-Log "Rollback/restart failed: $($_.Exception.Message)" 'ERROR'
        }
    }

    if ($TempRoot) {
        Write-Log "Troubleshooting files retained at $TempRoot." 'WARNING'
    }

    exit 1
}
