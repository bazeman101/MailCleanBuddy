<#
.SYNOPSIS
    Configuration Management module for MailCleanBuddy
.DESCRIPTION
    Manages application configuration with JSON storage and validation
#>

# Script-level configuration storage
$Script:Config = $null
$Script:ConfigFilePath = $null

<#
.SYNOPSIS
    Gets the default configuration
#>
function Get-DefaultConfiguration {
    return @{
        Version = "3.1"
        LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        # Logging settings
        Logging = @{
            Enabled = $true
            LogLevel = "Info"  # Error, Warning, Info, Debug
            EnableConsoleOutput = $false
            MaxLogSizeBytes = 5242880  # 5MB
            MaxLogFiles = 5
            LogRetentionDays = 30
        }

        # Cache settings
        Cache = @{
            AutoRefreshEnabled = $true
            AutoRefreshIntervalHours = 24
            MaxCacheAgeHours = 48
            EnableCacheValidation = $true
            MaxCacheSize = 100MB
        }

        # Email settings
        Email = @{
            MaxEmailsToIndex = 0  # 0 = unlimited
            DefaultPageSize = 30
            DefaultTimeFilter = "Last30Days"
            PreviewBodyLines = 30
        }

        # UI settings
        UI = @{
            ColorScheme = "Default"  # Default, Dark, Light
            UseEmojis = $true
            ShowNavigationHints = $true
            ConfirmDestructiveActions = $true
        }

        # Search settings
        Search = @{
            EnableFuzzySearch = $true
            FuzzySearchThreshold = 0.8
            SearchHistorySize = 20
            DefaultSearchScope = "All"  # Subject, Body, From, All
        }

        # Filter settings
        Filters = @{
            SavedFilters = @()
            DefaultFilters = @{
                HasAttachments = $null
                IsRead = $null
                Importance = $null
                MinSize = 0
                MaxSize = 0
            }
        }

        # Bulk Operations settings
        BulkOperations = @{
            EnableParallelProcessing = $true
            MaxParallelThreads = 4
            ShowProgressBar = $true
            BatchSize = 50
        }

        # Rules Engine settings
        Rules = @{
            Enabled = $true
            AutoExecuteRules = $false
            Rules = @()
        }

        # Performance settings
        Performance = @{
            EnableCaching = $true
            CacheTimeout = 300  # seconds
            MaxConcurrentApiCalls = 3
            ApiThrottleDelay = 100  # milliseconds
        }

        # Export settings
        Export = @{
            DefaultFormat = "EML"  # EML, MSG, CSV, JSON
            DefaultExportPath = ".\exports"
            IncludeAttachments = $true
            CompressExports = $false
        }
    }
}

<#
.SYNOPSIS
    Initializes configuration system
.PARAMETER ConfigDirectory
    Directory for config file (default: ~/.mailcleanbuddy)
#>
function Initialize-Configuration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ConfigDirectory
    )

    try {
        # Set config directory
        if ([string]::IsNullOrWhiteSpace($ConfigDirectory)) {
            $homeDir = if ($IsWindows -or $null -eq $IsWindows) { $env:USERPROFILE } else { $env:HOME }
            $ConfigDirectory = Join-Path $homeDir ".mailcleanbuddy"
        }

        # Create directory if needed
        if (-not (Test-Path $ConfigDirectory)) {
            New-Item -Path $ConfigDirectory -ItemType Directory -Force | Out-Null
        }

        # Set config file path
        $Script:ConfigFilePath = Join-Path $ConfigDirectory "config.json"

        # Load or create config
        if (Test-Path $Script:ConfigFilePath) {
            $Script:Config = Import-Configuration
            Write-Verbose "Configuration loaded from: $Script:ConfigFilePath"
        } else {
            $Script:Config = Get-DefaultConfiguration
            Export-Configuration
            Write-Verbose "Default configuration created at: $Script:ConfigFilePath"
        }

        return $true
    } catch {
        Write-Warning "Failed to initialize configuration: $($_.Exception.Message)"
        $Script:Config = Get-DefaultConfiguration
        return $false
    }
}

<#
.SYNOPSIS
    Imports configuration from file
#>
function Import-Configuration {
    [CmdletBinding()]
    param()

    try {
        if (-not (Test-Path $Script:ConfigFilePath)) {
            throw "Configuration file not found: $Script:ConfigFilePath"
        }

        $jsonContent = Get-Content -Path $Script:ConfigFilePath -Raw -Encoding UTF8
        $loadedConfig = $jsonContent | ConvertFrom-Json -AsHashtable

        # Merge with defaults to ensure all keys exist
        $defaultConfig = Get-DefaultConfiguration
        $mergedConfig = Merge-Configurations -Default $defaultConfig -Loaded $loadedConfig

        return $mergedConfig
    } catch {
        Write-Warning "Failed to import configuration: $($_.Exception.Message)"
        return Get-DefaultConfiguration
    }
}

<#
.SYNOPSIS
    Exports configuration to file
#>
function Export-Configuration {
    [CmdletBinding()]
    param()

    try {
        if (-not $Script:Config) {
            throw "Configuration not initialized"
        }

        # Update timestamp
        $Script:Config.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        # Convert to JSON
        $jsonContent = $Script:Config | ConvertTo-Json -Depth 10

        # Write to file
        Set-Content -Path $Script:ConfigFilePath -Value $jsonContent -Encoding UTF8

        Write-Verbose "Configuration saved to: $Script:ConfigFilePath"
        return $true
    } catch {
        Write-Warning "Failed to export configuration: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Gets a configuration value
.PARAMETER Path
    Configuration path (dot-notation, e.g., "Logging.LogLevel")
.PARAMETER DefaultValue
    Default value if path not found
#>
function Get-ConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        $DefaultValue = $null
    )

    try {
        if (-not $Script:Config) {
            Initialize-Configuration | Out-Null
        }

        $parts = $Path -split '\.'
        $current = $Script:Config

        foreach ($part in $parts) {
            if ($current -is [hashtable] -and $current.ContainsKey($part)) {
                $current = $current[$part]
            } else {
                return $DefaultValue
            }
        }

        return $current
    } catch {
        Write-Verbose "Error getting config value for path '$Path': $($_.Exception.Message)"
        return $DefaultValue
    }
}

<#
.SYNOPSIS
    Sets a configuration value
.PARAMETER Path
    Configuration path (dot-notation)
.PARAMETER Value
    Value to set
.PARAMETER SaveImmediately
    Save configuration to file immediately
#>
function Set-ConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        $Value,

        [Parameter(Mandatory = $false)]
        [switch]$SaveImmediately
    )

    try {
        if (-not $Script:Config) {
            Initialize-Configuration | Out-Null
        }

        $parts = $Path -split '\.'
        $current = $Script:Config

        # Navigate to parent
        for ($i = 0; $i -lt ($parts.Count - 1); $i++) {
            $part = $parts[$i]
            if (-not $current.ContainsKey($part)) {
                $current[$part] = @{}
            }
            $current = $current[$part]
        }

        # Set value
        $lastPart = $parts[-1]
        $current[$lastPart] = $Value

        Write-Verbose "Config value set: $Path = $Value"

        # Save if requested
        if ($SaveImmediately) {
            return Export-Configuration
        }

        return $true
    } catch {
        Write-Warning "Failed to set config value '$Path': $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Merges two configuration hashtables
#>
function Merge-Configurations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Default,

        [Parameter(Mandatory = $true)]
        [hashtable]$Loaded
    )

    $merged = $Default.Clone()

    foreach ($key in $Loaded.Keys) {
        if ($merged.ContainsKey($key)) {
            if ($merged[$key] -is [hashtable] -and $Loaded[$key] -is [hashtable]) {
                # Recursive merge for nested hashtables
                $merged[$key] = Merge-Configurations -Default $merged[$key] -Loaded $Loaded[$key]
            } else {
                # Direct assignment for non-hashtable values
                $merged[$key] = $Loaded[$key]
            }
        } else {
            # Add new key from loaded config
            $merged[$key] = $Loaded[$key]
        }
    }

    return $merged
}

<#
.SYNOPSIS
    Resets configuration to defaults
#>
function Reset-Configuration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$SaveImmediately
    )

    $Script:Config = Get-DefaultConfiguration

    if ($SaveImmediately) {
        return Export-Configuration
    }

    return $true
}

<#
.SYNOPSIS
    Gets the entire configuration
#>
function Get-Configuration {
    if (-not $Script:Config) {
        Initialize-Configuration | Out-Null
    }
    return $Script:Config.Clone()
}

<#
.SYNOPSIS
    Validates configuration integrity
#>
function Test-ConfigurationIntegrity {
    [CmdletBinding()]
    param()

    try {
        if (-not $Script:Config) {
            return $false
        }

        # Check required keys
        $requiredKeys = @("Version", "LastUpdated", "Logging", "Cache", "Email", "UI")

        foreach ($key in $requiredKeys) {
            if (-not $Script:Config.ContainsKey($key)) {
                Write-Warning "Missing required config key: $key"
                return $false
            }
        }

        # Validate specific values
        if ($Script:Config.Logging.MaxLogSizeBytes -lt 1MB) {
            Write-Warning "MaxLogSizeBytes too small (minimum 1MB)"
            return $false
        }

        if ($Script:Config.Cache.MaxCacheSize -lt 10MB) {
            Write-Warning "MaxCacheSize too small (minimum 10MB)"
            return $false
        }

        return $true
    } catch {
        Write-Warning "Configuration validation failed: $($_.Exception.Message)"
        return $false
    }
}

# Export functions
Export-ModuleMember -Function Initialize-Configuration, Import-Configuration, Export-Configuration, `
                              Get-ConfigValue, Set-ConfigValue, Get-Configuration, Reset-Configuration, `
                              Test-ConfigurationIntegrity, Get-DefaultConfiguration
