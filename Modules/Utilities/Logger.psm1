<#
.SYNOPSIS
    Logging module for MailCleanBuddy
.DESCRIPTION
    Provides comprehensive logging with multiple log levels, rotation, and export capabilities
#>

# Script-level variables
$Script:LogConfig = @{
    LogLevel = "Info"  # Error, Warning, Info, Debug
    LogFilePath = $null
    MaxLogSizeBytes = 5MB
    MaxLogFiles = 5
    EnableConsoleOutput = $false
    LogFormat = "[{0}] [{1}] {2}"  # Timestamp, Level, Message
}

<#
.SYNOPSIS
    Initializes the logging system
.PARAMETER LogDirectory
    Directory for log files (default: ~/.mailcleanbuddy/logs)
.PARAMETER LogLevel
    Minimum log level (Error, Warning, Info, Debug)
.PARAMETER EnableConsoleOutput
    Also output logs to console
#>
function Initialize-Logger {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$LogDirectory,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warning", "Info", "Debug")]
        [string]$LogLevel = "Info",

        [Parameter(Mandatory = $false)]
        [switch]$EnableConsoleOutput
    )

    try {
        # Set log directory
        if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
            $homeDir = if ($IsWindows -or $null -eq $IsWindows) { $env:USERPROFILE } else { $env:HOME }
            $LogDirectory = Join-Path $homeDir ".mailcleanbuddy" "logs"
        }

        # Create directory if needed
        if (-not (Test-Path $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }

        # Set log file path with date
        $logFileName = "MailCleanBuddy_$(Get-Date -Format 'yyyy-MM-dd').log"
        $Script:LogConfig.LogFilePath = Join-Path $LogDirectory $logFileName
        $Script:LogConfig.LogLevel = $LogLevel
        $Script:LogConfig.EnableConsoleOutput = $EnableConsoleOutput.IsPresent

        # Write initialization message
        Write-LogMessage -Level "Info" -Message "Logger initialized - LogLevel: $LogLevel, LogPath: $($Script:LogConfig.LogFilePath)"

        # Rotate old logs
        Invoke-LogRotation -LogDirectory $LogDirectory

        return $true
    } catch {
        Write-Warning "Failed to initialize logger: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Gets the numeric level value
#>
function Get-LogLevelValue {
    param([string]$Level)

    switch ($Level) {
        "Error"   { return 1 }
        "Warning" { return 2 }
        "Info"    { return 3 }
        "Debug"   { return 4 }
        default   { return 0 }
    }
}

<#
.SYNOPSIS
    Writes a log message
.PARAMETER Level
    Log level (Error, Warning, Info, Debug)
.PARAMETER Message
    Log message
.PARAMETER Exception
    Exception object (optional)
.PARAMETER Source
    Source module/function (optional)
#>
function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Error", "Warning", "Info", "Debug")]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [System.Exception]$Exception,

        [Parameter(Mandatory = $false)]
        [string]$Source
    )

    try {
        # Check if level is high enough to log
        $currentLevelValue = Get-LogLevelValue -Level $Script:LogConfig.LogLevel
        $messageLevelValue = Get-LogLevelValue -Level $Level

        if ($messageLevelValue -gt $currentLevelValue) {
            return  # Don't log if message level is lower than configured level
        }

        # Format timestamp
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"

        # Build log message
        $logMessage = $Script:LogConfig.LogFormat -f $timestamp, $Level.ToUpper(), $Message

        # Add source if provided
        if (-not [string]::IsNullOrWhiteSpace($Source)) {
            $logMessage += " [Source: $Source]"
        }

        # Add exception details if provided
        if ($Exception) {
            $logMessage += "`n  Exception: $($Exception.Message)"
            if ($Exception.StackTrace) {
                $logMessage += "`n  StackTrace: $($Exception.StackTrace)"
            }
        }

        # Write to file
        if ($Script:LogConfig.LogFilePath) {
            # Check log size and rotate if needed
            if (Test-Path $Script:LogConfig.LogFilePath) {
                $logFile = Get-Item $Script:LogConfig.LogFilePath
                if ($logFile.Length -gt $Script:LogConfig.MaxLogSizeBytes) {
                    Invoke-LogRotation -LogDirectory (Split-Path $Script:LogConfig.LogFilePath -Parent)
                }
            }

            Add-Content -Path $Script:LogConfig.LogFilePath -Value $logMessage -Encoding UTF8
        }

        # Write to console if enabled
        if ($Script:LogConfig.EnableConsoleOutput) {
            $color = switch ($Level) {
                "Error"   { "Red" }
                "Warning" { "Yellow" }
                "Info"    { "Cyan" }
                "Debug"   { "Gray" }
                default   { "White" }
            }
            Write-Host $logMessage -ForegroundColor $color
        }
    } catch {
        # Fallback: write to console if logging fails
        Write-Warning "Logging failed: $($_.Exception.Message) - Original message: $Message"
    }
}

<#
.SYNOPSIS
    Rotates log files
#>
function Invoke-LogRotation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogDirectory
    )

    try {
        # Get all log files sorted by creation time (oldest first)
        $logFiles = Get-ChildItem -Path $LogDirectory -Filter "MailCleanBuddy_*.log" -File |
                    Sort-Object CreationTime

        # If we have more files than max, delete oldest
        $maxFiles = $Script:LogConfig.MaxLogFiles
        if ($logFiles.Count -gt $maxFiles) {
            $filesToDelete = $logFiles | Select-Object -First ($logFiles.Count - $maxFiles)
            foreach ($file in $filesToDelete) {
                Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
                Write-LogMessage -Level "Info" -Message "Rotated old log file: $($file.Name)"
            }
        }

        # Also delete log files older than 30 days
        $cutoffDate = (Get-Date).AddDays(-30)
        $oldFiles = $logFiles | Where-Object { $_.CreationTime -lt $cutoffDate }
        foreach ($file in $oldFiles) {
            Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
            Write-LogMessage -Level "Info" -Message "Deleted old log file: $($file.Name)"
        }
    } catch {
        Write-Warning "Log rotation failed: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Exports logs matching filter criteria
.PARAMETER OutputPath
    Output file path
.PARAMETER StartDate
    Start date filter
.PARAMETER EndDate
    End date filter
.PARAMETER LevelFilter
    Log level filter
#>
function Export-Logs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [DateTime]$StartDate,

        [Parameter(Mandatory = $false)]
        [DateTime]$EndDate,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warning", "Info", "Debug")]
        [string]$LevelFilter
    )

    try {
        if (-not $Script:LogConfig.LogFilePath) {
            Write-Warning "Logger not initialized"
            return $false
        }

        $logDir = Split-Path $Script:LogConfig.LogFilePath -Parent
        $logFiles = Get-ChildItem -Path $logDir -Filter "MailCleanBuddy_*.log" -File

        $filteredLines = @()

        foreach ($logFile in $logFiles) {
            # Check date range if provided
            if ($StartDate -and $logFile.CreationTime -lt $StartDate) { continue }
            if ($EndDate -and $logFile.CreationTime -gt $EndDate) { continue }

            $content = Get-Content -Path $logFile.FullName

            foreach ($line in $content) {
                # Apply level filter if provided
                if ($LevelFilter -and $line -notmatch "\[$LevelFilter\]") { continue }

                $filteredLines += $line
            }
        }

        # Write to output file
        $filteredLines | Set-Content -Path $OutputPath -Encoding UTF8

        Write-LogMessage -Level "Info" -Message "Exported $($filteredLines.Count) log entries to: $OutputPath"
        return $true
    } catch {
        Write-LogMessage -Level "Error" -Message "Failed to export logs" -Exception $_.Exception
        return $false
    }
}

<#
.SYNOPSIS
    Gets the current log configuration
#>
function Get-LoggerConfiguration {
    return $Script:LogConfig.Clone()
}

<#
.SYNOPSIS
    Sets log level dynamically
#>
function Set-LogLevel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Error", "Warning", "Info", "Debug")]
        [string]$LogLevel
    )

    $Script:LogConfig.LogLevel = $LogLevel
    Write-LogMessage -Level "Info" -Message "Log level changed to: $LogLevel"
}

# Export functions
Export-ModuleMember -Function Initialize-Logger, Write-LogMessage, Export-Logs, Get-LoggerConfiguration, Set-LogLevel, Invoke-LogRotation
