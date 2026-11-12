[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$AppId,

    [ValidateSet("InstallOrUpdate","Uninstall")]
    [string]$Mode = "InstallOrUpdate",

    [ValidateSet("Machine","User","Default")]
    [string]$Scope = "Machine",

    # Most applications use winget.
    # Microsoft Store applications can use msstore.
    [ValidateSet("winget","msstore")]
    [string]$Source = "winget",

    # Optional installer type.
    # Example for PowerShell 7 MSI/WiX:
    # -InstallerType wix
    [string]$InstallerType,

    # Optional physical path used to verify the required installation exists.
    # Particularly useful where multiple installer technologies exist for
    # the same Winget AppId.
    #
    # Example:
    # C:\Program Files\PowerShell\7\pwsh.exe
    [string]$DetectionPath,

    # Force an installation attempt
    [switch]$ForceReinstall,

    # If upgrade/forced install fails, optionally uninstall then install
    [switch]$UninstallThenInstall,

    # Reset Winget sources before operation
    [switch]$ResetSources,

    # Log directory
    [string]$LogRoot = "C:\ProgramData\customapps",

    # Maximum time to wait for Winget operations
    [int]$TimeoutSeconds = 600,

    # Number of retry attempts after the first attempt
    [int]$MaxRetries = 2
)

# ============================================================
# Input validation
# ============================================================

if ([string]::IsNullOrWhiteSpace($AppId)) {
    Write-Error "AppId cannot be empty or whitespace."
    exit 1
}

if ($AppId -notmatch '^[a-zA-Z0-9._-]+$') {
    Write-Warning "AppId '$AppId' contains unusual characters."
}

# ============================================================
# Logging
# ============================================================

try {
    New-Item `
        -ItemType Directory `
        -Path $LogRoot `
        -Force `
        -ErrorAction Stop |
        Out-Null
}
catch {
    Write-Error "Failed to create log directory '$LogRoot': $($_.Exception.Message)"
    exit 1
}

$SafeAppId = $AppId -replace '[\\/:*?"<>|]', '_'

$ScriptLog = Join-Path `
    $LogRoot `
    ("winget-wrapper-{0}.log" -f $SafeAppId)

$WingetLog = Join-Path `
    $LogRoot `
    ("winget-{0}-{1}.log" -f $SafeAppId, (Get-Date -Format "yyyyMMdd-HHmmss"))

function Write-Log {

    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO",

        [switch]$SuppressConsole
    )

    $Line = "[{0}] [{1}] {2}" -f `
        (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
        $Level,
        $Message

    try {
        $Line |
            Out-File `
                -FilePath $ScriptLog `
                -Append `
                -Encoding utf8 `
                -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to write log entry: $($_.Exception.Message)"
    }

    if (-not $SuppressConsole) {

        switch ($Level) {

            "ERROR" {
                Write-Host $Line -ForegroundColor Red
            }

            "WARN" {
                Write-Host $Line -ForegroundColor Yellow
            }

            default {
                Write-Verbose $Line
            }
        }
    }
}

# ============================================================
# Find Winget
# ============================================================

function Get-WingetPath {

    # First try PATH
    $Command = Get-Command winget.exe -ErrorAction SilentlyContinue

    if ($Command -and $Command.Source) {

        Write-Log "Found winget through PATH: $($Command.Source)"

        return $Command.Source
    }

    # Then find Desktop App Installer
    try {

        $Package = Get-AppxPackage `
            -AllUsers `
            -Name "Microsoft.DesktopAppInstaller" `
            -ErrorAction Stop |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($Package -and $Package.InstallLocation) {

            $Executable = Join-Path `
                $Package.InstallLocation `
                "winget.exe"

            if (Test-Path -LiteralPath $Executable) {

                Write-Log "Found winget through Desktop App Installer: $Executable"

                return $Executable
            }
        }
    }
    catch {
        Write-Log `
            "Failed to query Desktop App Installer: $($_.Exception.Message)" `
            "WARN"
    }

    # Fallback locations
    $CommonPaths = @(
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
        "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*\winget.exe"
    )

    foreach ($Path in $CommonPaths) {

        $Resolved = Get-Item `
            $Path `
            -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            Select-Object -First 1

        if ($Resolved) {

            Write-Log "Found winget at: $($Resolved.FullName)"

            return $Resolved.FullName
        }
    }

    return $null
}

# ============================================================
# SYSTEM safety
# ============================================================

if (
    $env:USERNAME -eq "SYSTEM" -and
    $Scope -eq "User"
) {

    Write-Log `
        "Script is running as SYSTEM. Changing requested scope from User to Machine." `
        "WARN"

    $Scope = "Machine"
}

# ============================================================
# Start logging
# ============================================================

Write-Log "============================================================"
Write-Log "Starting Winget wrapper"
Write-Log "AppId=$AppId"
Write-Log "Mode=$Mode"
Write-Log "Scope=$Scope"
Write-Log "Source=$Source"
Write-Log "InstallerType=$InstallerType"
Write-Log "DetectionPath=$DetectionPath"
Write-Log "ForceReinstall=$ForceReinstall"
Write-Log "UninstallThenInstall=$UninstallThenInstall"
Write-Log "User=$env:USERNAME"
Write-Log "Is64BitProcess=$([Environment]::Is64BitProcess)"
Write-Log "TimeoutSeconds=$TimeoutSeconds"
Write-Log "MaxRetries=$MaxRetries"

# ============================================================
# Locate Winget
# ============================================================

$Winget = Get-WingetPath

if (-not $Winget) {

    Write-Log `
        "winget.exe could not be located." `
        "ERROR"

    exit 1
}

# ============================================================
# Prepare WinGet AppX dependencies for SYSTEM
# ============================================================

$IsSystem = [Security.Principal.WindowsIdentity]::GetCurrent().IsSystem

Write-Log "IsSystem=$IsSystem"

if ($IsSystem) {
    Write-Log "SYSTEM context detected. Preparing WinGet AppX dependencies."
    $WingetDependencyPackages = @(
        "Microsoft.WindowsAppRuntime.1.8",
        "Microsoft.VCLibs.140.00",
        "Microsoft.VCLibs.140.00.UWPDesktop"
    )
    foreach ($DependencyName in $WingetDependencyPackages) {
        try {
            $Dependency = Get-AppxPackage -AllUsers -Name $DependencyName -ErrorAction Stop |
                Where-Object {
                    $_.Architecture -eq "X64" -and
                    -not [string]::IsNullOrWhiteSpace($_.InstallLocation) -and
                    (Test-Path -LiteralPath $_.InstallLocation)
                } |
                Sort-Object Version -Descending |
                Select-Object -First 1
            if ($Dependency) {
                Write-Log "Found WinGet dependency: $($Dependency.Name) $($Dependency.Version)"
                Write-Log "Adding WinGet dependency to process PATH: $($Dependency.InstallLocation)"
                $ExistingPaths = $env:PATH -split ';'
                if ($ExistingPaths -notcontains $Dependency.InstallLocation) {
                    $env:PATH = "$($Dependency.InstallLocation);$env:PATH"
                }
            } else {
                Write-Log "WinGet dependency not found: $DependencyName" "WARN"
            }
        } catch {
            Write-Log "Failed to query WinGet dependency '$DependencyName': $($_.Exception.Message)" "WARN"
        }
    }
}

# ============================================================
# Validate Winget
# ============================================================

try {

    $VersionOutput = & $Winget --version 2>&1

    $VersionString = ($VersionOutput | Out-String).Trim()

    if (-not [string]::IsNullOrWhiteSpace($VersionString)) {

        Write-Log "Winget version: $VersionString"
    }
    else {

        Write-Log `
            "Winget version output was empty. Continuing because winget.exe was successfully located." `
            "WARN"
    }
}
catch {

    Write-Log `
        "Unable to retrieve Winget version: $($_.Exception.Message). Continuing because winget.exe was successfully located." `
        "WARN"
}

# ============================================================
# Source reset
# ============================================================

if ($ResetSources) {

    try {

        Write-Log "Resetting Winget sources..."

        $ResetOutput = & $Winget `
            source reset `
            --force `
            2>&1

        Write-Log "Source reset output:`n$($ResetOutput | Out-String)"

        Write-Log "Updating Winget sources..."

        $UpdateOutput = & $Winget `
            source update `
            2>&1

        Write-Log "Source update output:`n$($UpdateOutput | Out-String)"
    }
    catch {

        Write-Log `
            "Source reset/update failed: $($_.Exception.Message)" `
            "WARN"
    }
}

# ============================================================
# Invoke Winget and capture output
# ============================================================

function Invoke-WingetCapture {

    param(
        [Parameter(Mandatory)]
        [string[]]$Args,

        [int]$Timeout = $TimeoutSeconds
    )

    $CommandLine = "$Winget $($Args -join ' ')"

    Write-Log "Executing: $CommandLine"

    try {

        $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo

        $ProcessInfo.FileName = $Winget
        $ProcessInfo.Arguments = $Args -join " "
        $ProcessInfo.RedirectStandardOutput = $true
        $ProcessInfo.RedirectStandardError = $true
        $ProcessInfo.UseShellExecute = $false
        $ProcessInfo.CreateNoWindow = $true

        $Process = New-Object System.Diagnostics.Process

        $Process.StartInfo = $ProcessInfo

        [void]$Process.Start()

        # Read asynchronously to avoid output-buffer deadlocks
        $StdOutTask = $Process.StandardOutput.ReadToEndAsync()
        $StdErrTask = $Process.StandardError.ReadToEndAsync()

        $Completed = $Process.WaitForExit($Timeout * 1000)

        if (-not $Completed) {

            Write-Log `
                "Winget operation exceeded timeout of $Timeout seconds." `
                "ERROR"

            try {
                $Process.Kill()
                $Process.WaitForExit(5000) | Out-Null
            }
            catch {
                Write-Log `
                    "Failed to terminate timed-out Winget process: $($_.Exception.Message)" `
                    "WARN"
            }

            return [pscustomobject]@{
                ExitCode = -1
                Output   = "Winget operation timed out after $Timeout seconds."
                TimedOut = $true
            }
        }

        $Process.WaitForExit()

        $StdOut = $StdOutTask.Result
        $StdErr = $StdErrTask.Result

        $Output = @(
            $StdOut
            $StdErr
        ) -join [Environment]::NewLine

        $ExitCode = $Process.ExitCode

        if (-not [string]::IsNullOrWhiteSpace($Output)) {

            Write-Log "Winget output:`n$Output"

            $Output |
                Out-File `
                    -FilePath $WingetLog `
                    -Append `
                    -Encoding utf8
        }

        Write-Log "Winget exit code: $ExitCode"

        return [pscustomobject]@{
            ExitCode = $ExitCode
            Output   = $Output.Trim()
            TimedOut = $false
        }
    }
    catch {

        Write-Log `
            "Exception executing Winget: $($_.Exception.Message)" `
            "ERROR"

        return [pscustomobject]@{
            ExitCode = 1
            Output   = $_.Exception.Message
            TimedOut = $false
        }
    }
}

# ============================================================
# Retry Winget
# ============================================================

function Invoke-WingetWithRetry {

    param(
        [Parameter(Mandatory)]
        [string[]]$Args,

        [int]$MaxAttempts = $MaxRetries
    )

    $Attempt = 0
    $LastResult = $null

    while ($Attempt -le $MaxAttempts) {

        $Attempt++

        if ($Attempt -gt 1) {

            $WaitSeconds = 5 * $Attempt

            Write-Log `
                "Retry $Attempt of $($MaxAttempts + 1). Waiting $WaitSeconds seconds..." `
                "WARN"

            Start-Sleep -Seconds $WaitSeconds
        }

        $LastResult = Invoke-WingetCapture -Args $Args

        if ($LastResult.ExitCode -eq 0) {

            return $LastResult
        }

        if (
            $LastResult.Output -match `
            "No available upgrade|No applicable upgrade|No newer package versions"
        ) {

            return $LastResult
        }

        if ($LastResult.TimedOut) {

            return $LastResult
        }

        $RetriablePatterns = @(
            "0x80070490",
            "0x80072ee7",
            "0x80072ef3",
            "0x80072efd",
            "Unable to connect",
            "network error",
            "Failed to open source",
            "source agreement"
        )

        $ShouldRetry = $false

        foreach ($Pattern in $RetriablePatterns) {

            if ($LastResult.Output -match [regex]::Escape($Pattern)) {

                $ShouldRetry = $true

                Write-Log `
                    "Detected potentially transient Winget failure: $Pattern" `
                    "WARN"

                break
            }
        }

        if (-not $ShouldRetry) {
            return $LastResult
        }
    }

    return $LastResult
}

# ============================================================
# Winget package detection
# ============================================================

function Test-WingetInstalled {

    param(
        [string]$CheckScope
    )

    $ScopeDisplay = if ($CheckScope) {
        $CheckScope
    }
    else {
        "Any"
    }

    Write-Log `
        "Checking Winget inventory for $AppId in scope: $ScopeDisplay"

    $Arguments = @(
        "list",
        "--id", $AppId,
        "--exact",
        "--accept-source-agreements",
        "--disable-interactivity"
    )

    if ($CheckScope) {

        $Arguments += @(
            "--scope",
            $CheckScope.ToLower()
        )
    }

    $Result = Invoke-WingetCapture `
        -Args $Arguments

    if (
        $Result.Output -match `
        "No installed package found|No installed packages? found matching|No package found"
    ) {

        Write-Log `
            "Winget does not report $AppId as installed in scope $ScopeDisplay."

        return $false
    }

    $EscapedAppId = [regex]::Escape($AppId)

    if ($Result.Output -match "(?m)^.*$EscapedAppId.*$") {

        Write-Log `
            "Winget reports $AppId as installed in scope $ScopeDisplay."

        return $true
    }

    Write-Log `
        "Unable to positively identify $AppId in Winget inventory for scope $ScopeDisplay." `
        "WARN"

    return $false
}

# ============================================================
# Physical detection path
# ============================================================

function Test-DetectionPath {

    if ([string]::IsNullOrWhiteSpace($DetectionPath)) {

        return $null
    }

    Write-Log "Testing detection path: $DetectionPath"

    if (Test-Path -LiteralPath $DetectionPath) {

        Write-Log "Detection path exists."

        return $true
    }

    Write-Log "Detection path does not exist."

    return $false
}

# ============================================================
# Determine current state
# ============================================================

function Test-RequiredInstallation {

    # If a physical detection path was supplied, it takes priority.
    #
    # This is important for applications like PowerShell where the same
    # Winget AppId can represent MSIX and WiX/MSI installations.

    if (-not [string]::IsNullOrWhiteSpace($DetectionPath)) {

        $PathResult = Test-DetectionPath

        Write-Log `
            "Required installation status based on DetectionPath: $PathResult"

        return $PathResult
    }

    # Otherwise use normal Winget detection.
    #
    # Scope "Default" means Winget should determine the appropriate
    # installer scope. In this case, do not pass --scope.

    $CheckScope = if ($Scope -eq "Default") {
        $null
    }
    else {
        $Scope
    }

    $WingetResult = Test-WingetInstalled `
        -CheckScope $CheckScope

    if ($Scope -eq "Default") {

        Write-Log `
            "Required installation status based on Winget default scope: $WingetResult"
    }
    else {

        Write-Log `
            "Required installation status based on Winget scope '$Scope': $WingetResult"
    }

    return $WingetResult
}

# ============================================================
# Common installation arguments
# ============================================================

$CommonArguments = @(
    "--id", $AppId,
    "--exact",
    "--source", $Source,
    "--accept-package-agreements",
    "--accept-source-agreements",
    "--disable-interactivity",
    "--silent"
)

# Only explicitly specify a Winget scope when Machine or User
# has been requested. Default leaves installer selection to Winget.

if ($Scope -ne "Default") {

    $CommonArguments += @(
        "--scope",
        $Scope.ToLower()
    )

    Write-Log "Winget scope explicitly requested: $Scope"
}
else {

    Write-Log `
        "Winget scope set to Default. No --scope argument will be passed."
}

if (-not [string]::IsNullOrWhiteSpace($InstallerType)) {

    $CommonArguments += @(
        "--installer-type",
        $InstallerType
    )

    Write-Log `
        "Installer type explicitly requested: $InstallerType"
}

# ============================================================
# Uninstall
# ============================================================

function Invoke-WingetUninstall {

    Write-Log "Attempting to uninstall $AppId..."

    $Arguments = @(
        "uninstall",
        "--id", $AppId,
        "--exact",
        "--source", $Source,
        "--accept-source-agreements",
        "--disable-interactivity",
        "--silent"
    )

    if ($Scope -ne "Default") {

        $Arguments += @(
            "--scope",
            $Scope.ToLower()
        )
    }

    $Result = Invoke-WingetWithRetry `
        -Args $Arguments

    return $Result
}

# ============================================================
# Execute operation
# ============================================================

$ReturnCode = 0

try {

    $Installed = Test-RequiredInstallation

    Write-Log `
        "Initial required installation state: $Installed"

    switch ($Mode) {

        # ----------------------------------------------------
        # INSTALL
        # ----------------------------------------------------

        "Install" {

            Write-Log "=== Mode: Install ==="

            if ($Installed -and -not $ForceReinstall) {

                Write-Log `
                    "Required installation is already present. No installation required."

                $ReturnCode = 0
            }
            else {

                $InstallArguments = @(
                    "install"
                ) + $CommonArguments

                if ($ForceReinstall) {

                    $InstallArguments += "--force"

                    Write-Log "Force reinstall requested."
                }

                $Result = Invoke-WingetWithRetry `
                    -Args $InstallArguments

                $ReturnCode = $Result.ExitCode
            }
        }

        # ----------------------------------------------------
        # UPGRADE
        # ----------------------------------------------------

        "Upgrade" {

            Write-Log "=== Mode: Upgrade ==="

            if (-not $Installed) {

                Write-Log `
                    "Application is not installed. Upgrade cannot be performed." `
                    "WARN"

                $ReturnCode = 0
            }
            else {

                $UpgradeArguments = @(
                    "upgrade"
                ) + $CommonArguments

                $Result = Invoke-WingetWithRetry `
                    -Args $UpgradeArguments

                if (
                    $Result.ExitCode -eq 0 -or
                    $Result.Output -match `
                    "No available upgrade|No applicable upgrade|No newer package versions"
                ) {

                    $ReturnCode = 0
                }
                else {

                    $ReturnCode = $Result.ExitCode
                }
            }
        }

        # ----------------------------------------------------
        # INSTALL OR UPDATE
        # ----------------------------------------------------

        "InstallOrUpdate" {

            Write-Log "=== Mode: InstallOrUpdate ==="

            if (-not $Installed) {

                Write-Log `
                    "Required installation was not detected. Performing fresh installation."

                $InstallArguments = @(
                    "install"
                ) + $CommonArguments

                # If an explicit InstallerType is requested, use --force.
                #
                # This prevents an existing alternative package type
                # such as MSIX from stopping installation of the requested
                # WiX/MSI package.

                if (
                    -not [string]::IsNullOrWhiteSpace($InstallerType) -or
                    $ForceReinstall
                ) {

                    $InstallArguments += "--force"

                    Write-Log `
                        "Using --force because an explicit installer type or force reinstall was requested."
                }

                $Result = Invoke-WingetWithRetry `
                    -Args $InstallArguments

                $ReturnCode = $Result.ExitCode
            }
            else {

                Write-Log `
                    "Required installation is already present. Checking for upgrade."

                $UpgradeArguments = @(
                    "upgrade"
                ) + $CommonArguments

                $UpgradeResult = Invoke-WingetWithRetry `
                    -Args $UpgradeArguments

                if (
                    $UpgradeResult.ExitCode -eq 0 -or
                    $UpgradeResult.Output -match `
                    "No available upgrade|No applicable upgrade|No newer package versions"
                ) {

                    Write-Log `
                        "Application is already current or upgrade completed successfully."

                    $ReturnCode = 0
                }
                else {

                    Write-Log `
                        "Upgrade failed. Attempting forced installation." `
                        "WARN"

                    $InstallArguments = @(
                        "install"
                    ) + $CommonArguments + @(
                        "--force"
                    )

                    $InstallResult = Invoke-WingetWithRetry `
                        -Args $InstallArguments

                    $ReturnCode = $InstallResult.ExitCode

                    if (
                        $ReturnCode -ne 0 -and
                        $UninstallThenInstall
                    ) {

                        Write-Log `
                            "Forced installation failed. Attempting uninstall and reinstall." `
                            "WARN"

                        $UninstallResult = Invoke-WingetUninstall

                        if ($UninstallResult.ExitCode -eq 0) {

                            Start-Sleep -Seconds 3

                            $ReinstallArguments = @(
                                "install"
                            ) + $CommonArguments

                            $ReinstallResult = Invoke-WingetWithRetry `
                                -Args $ReinstallArguments

                            $ReturnCode = $ReinstallResult.ExitCode
                        }
                        else {

                            Write-Log `
                                "Uninstall failed with exit code $($UninstallResult.ExitCode)." `
                                "ERROR"

                            $ReturnCode = $UninstallResult.ExitCode
                        }
                    }
                }
            }
        }

        # ----------------------------------------------------
        # UNINSTALL
        # ----------------------------------------------------

        "Uninstall" {

            Write-Log "=== Mode: Uninstall ==="

            if (-not $Installed) {

                Write-Log `
                    "Application is not currently detected. Nothing to uninstall."

                $ReturnCode = 0
            }
            else {

                $UninstallResult = Invoke-WingetUninstall

                $ReturnCode = $UninstallResult.ExitCode

                if ($ReturnCode -eq 0) {

                    Write-Log `
                        "Winget reported that the uninstall completed successfully."
                }
                else {

                    Write-Log `
                        "Winget uninstall failed with exit code $ReturnCode." `
                        "ERROR"
                }
            }
        }
    }
}
catch {

    Write-Log `
        "Unhandled exception: $($_.Exception.Message)" `
        "ERROR"

    Write-Log `
        "Stack trace: $($_.ScriptStackTrace)" `
        "ERROR"

    $ReturnCode = 1
}

# ============================================================
# Final verification
# ============================================================

Write-Log "============================================================"
Write-Log "Performing final verification..."
Write-Log "Winget operation exit code: $ReturnCode"

try {

    $FinalInstalled = Test-RequiredInstallation

    Write-Log `
        "Final required installation state: $FinalInstalled"

    # --------------------------------------------------------
    # UNINSTALL VERIFICATION
    # --------------------------------------------------------

    if ($Mode -eq "Uninstall") {

        if ($FinalInstalled) {

            Write-Log `
                "Application is still detected after uninstall." `
                "ERROR"

            $ReturnCode = 1
        }
        else {

            Write-Log `
                "Application is no longer detected. Uninstall verified successfully."

            $ReturnCode = 0
        }
    }

    # --------------------------------------------------------
    # INSTALL / INSTALLORUPDATE VERIFICATION
    # --------------------------------------------------------

    elseif ($Mode -ne "Upgrade") {

        if (-not $FinalInstalled) {

            Write-Log `
                "Required installation was not detected after Winget completed." `
                "ERROR"

            $ReturnCode = 1
        }
        else {

            Write-Log `
                "Required installation detected successfully."
        }
    }
}
catch {

    Write-Log `
        "Final verification failed: $($_.Exception.Message)" `
        "ERROR"

    $ReturnCode = 1
}

# ============================================================
# Final status
# ============================================================

Write-Log "Winget detailed log: $WingetLog"

if ($ReturnCode -eq 0) {

    Write-Log `
        "SUCCESS: Operation completed successfully."
}
else {

    Write-Log `
        "FAILURE: Operation completed with exit code $ReturnCode." `
        "ERROR"
}

Write-Log `
    "Script finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

Write-Log "============================================================"

exit $ReturnCode
