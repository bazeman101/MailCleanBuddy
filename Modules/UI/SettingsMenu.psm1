<#
.SYNOPSIS
    Settings and Configuration Menu for MailCleanBuddy
.DESCRIPTION
    Provides interactive menu for viewing and modifying configuration settings
#>

<#
.SYNOPSIS
    Shows the settings configuration menu
#>
function Show-SettingsMenu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    $continue = $true
    while ($continue) {
        Clear-Host
        Write-Host "`n⚙️  MailCleanBuddy Settings" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Load current config
        $config = Get-Configuration

        # Display current settings
        Write-Host "📊 Current Settings:" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ""

        Write-Host "  Cache Settings:" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "    Max Cache Age: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Cache.MaxCacheAgeHours) hours" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "    Auto Refresh: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Cache.AutoRefreshEnabled)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "    Auto Refresh Interval: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Cache.AutoRefreshIntervalHours) hours" -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""

        Write-Host "  Email Display Settings:" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "    Default Page Size: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Email.DefaultPageSize) emails per page" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "    Max Emails to Index: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        $maxIndex = if ($config.Email.MaxEmailsToIndex -eq 0) { "Unlimited" } else { $config.Email.MaxEmailsToIndex }
        Write-Host "$maxIndex" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "    Preview Body Lines: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Email.PreviewBodyLines) lines" -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""

        Write-Host "  UI Settings:" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "    Use Emojis: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.UI.UseEmojis)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "    Color Scheme: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.UI.ColorScheme)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""

        Write-Host "  Logging Settings:" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "    Enabled: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Logging.Enabled)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "    Log Level: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Logging.LogLevel)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "    Max Log Files: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Logging.MaxLogFiles)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""

        Write-Host "  Search Settings:" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "    Fuzzy Search: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Search.EnableFuzzySearch)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "    Fuzzy Threshold: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Search.FuzzySearchThreshold)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""

        Write-Host "  Rule Engine Settings:" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "    Enabled: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Rules.Enabled)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "    Auto Execute: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($config.Rules.AutoExecuteRules)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""

        Write-Host "⚙️  Configuration Actions:" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  [1] Change Cache Settings" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  [2] Change Email Display Settings" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  [3] Change UI Settings" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  [4] Change Logging Settings" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  [5] Change Search Settings" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  [6] Change Rule Engine Settings" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  [R] Reset to Defaults" -ForegroundColor $Global:ColorScheme.Warning
        Write-Host "  [Q] Back to Main Menu" -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        $choice = Read-Host "Select option"

        switch ($choice.ToUpper()) {
            "1" {
                Edit-CacheSettings
            }
            "2" {
                Edit-EmailDisplaySettings
            }
            "3" {
                Edit-UISettings
            }
            "4" {
                Edit-LoggingSettings
            }
            "5" {
                Edit-SearchSettings
            }
            "6" {
                Edit-RuleEngineSettings
            }
            "R" {
                $confirm = Read-Host "Reset all settings to defaults? (yes/no)"
                if ($confirm -eq "yes" -or $confirm -eq "y") {
                    Reset-Configuration -SaveImmediately
                    Write-Host "Settings reset to defaults!" -ForegroundColor $Global:ColorScheme.Success
                    Start-Sleep -Seconds 2
                }
            }
            "Q" {
                $continue = $false
            }
            default {
                Write-Host "Invalid option." -ForegroundColor $Global:ColorScheme.Warning
                Start-Sleep -Seconds 1
            }
        }
    }
}

<#
.SYNOPSIS
    Edits cache settings
#>
function Edit-CacheSettings {
    Clear-Host
    Write-Host "`n⚙️  Cache Settings" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    $maxAge = Read-Host "Max Cache Age (hours, current: $(Get-ConfigValue 'Cache.MaxCacheAgeHours'))"
    if ($maxAge -match '^\d+$') {
        Set-ConfigValue -Path "Cache.MaxCacheAgeHours" -Value ([int]$maxAge) -SaveImmediately
    }

    $autoRefresh = Read-Host "Auto Refresh Enabled (true/false, current: $(Get-ConfigValue 'Cache.AutoRefreshEnabled'))"
    if ($autoRefresh -in @("true", "false")) {
        Set-ConfigValue -Path "Cache.AutoRefreshEnabled" -Value ([bool]::Parse($autoRefresh)) -SaveImmediately
    }

    $refreshInterval = Read-Host "Auto Refresh Interval (hours, current: $(Get-ConfigValue 'Cache.AutoRefreshIntervalHours'))"
    if ($refreshInterval -match '^\d+$') {
        Set-ConfigValue -Path "Cache.AutoRefreshIntervalHours" -Value ([int]$refreshInterval) -SaveImmediately
    }

    Write-Host "Cache settings updated!" -ForegroundColor $Global:ColorScheme.Success
    Start-Sleep -Seconds 2
}

<#
.SYNOPSIS
    Edits email display settings
#>
function Edit-EmailDisplaySettings {
    Clear-Host
    Write-Host "`n⚙️  Email Display Settings" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    $pageSize = Read-Host "Default Page Size (emails per page, current: $(Get-ConfigValue 'Email.DefaultPageSize'))"
    if ($pageSize -match '^\d+$' -and [int]$pageSize -gt 0) {
        Set-ConfigValue -Path "Email.DefaultPageSize" -Value ([int]$pageSize) -SaveImmediately
    }

    $maxIndex = Read-Host "Max Emails to Index (0=unlimited, current: $(Get-ConfigValue 'Email.MaxEmailsToIndex'))"
    if ($maxIndex -match '^\d+$') {
        Set-ConfigValue -Path "Email.MaxEmailsToIndex" -Value ([int]$maxIndex) -SaveImmediately
    }

    $previewLines = Read-Host "Preview Body Lines (current: $(Get-ConfigValue 'Email.PreviewBodyLines'))"
    if ($previewLines -match '^\d+$' -and [int]$previewLines -gt 0) {
        Set-ConfigValue -Path "Email.PreviewBodyLines" -Value ([int]$previewLines) -SaveImmediately
    }

    Write-Host "Email display settings updated!" -ForegroundColor $Global:ColorScheme.Success
    Start-Sleep -Seconds 2
}

<#
.SYNOPSIS
    Edits UI settings
#>
function Edit-UISettings {
    Clear-Host
    Write-Host "`n⚙️  UI Settings" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    $useEmojis = Read-Host "Use Emojis (true/false, current: $(Get-ConfigValue 'UI.UseEmojis'))"
    if ($useEmojis -in @("true", "false")) {
        Set-ConfigValue -Path "UI.UseEmojis" -Value ([bool]::Parse($useEmojis)) -SaveImmediately
    }

    Write-Host "UI settings updated!" -ForegroundColor $Global:ColorScheme.Success
    Start-Sleep -Seconds 2
}

<#
.SYNOPSIS
    Edits logging settings
#>
function Edit-LoggingSettings {
    Clear-Host
    Write-Host "`n⚙️  Logging Settings" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    $enabled = Read-Host "Logging Enabled (true/false, current: $(Get-ConfigValue 'Logging.Enabled'))"
    if ($enabled -in @("true", "false")) {
        Set-ConfigValue -Path "Logging.Enabled" -Value ([bool]::Parse($enabled)) -SaveImmediately
    }

    $logLevel = Read-Host "Log Level (Error/Warning/Info/Debug, current: $(Get-ConfigValue 'Logging.LogLevel'))"
    if ($logLevel -in @("Error", "Warning", "Info", "Debug")) {
        Set-ConfigValue -Path "Logging.LogLevel" -Value $logLevel -SaveImmediately
    }

    Write-Host "Logging settings updated!" -ForegroundColor $Global:ColorScheme.Success
    Start-Sleep -Seconds 2
}

<#
.SYNOPSIS
    Edits search settings
#>
function Edit-SearchSettings {
    Clear-Host
    Write-Host "`n⚙️  Search Settings" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    $fuzzyEnabled = Read-Host "Fuzzy Search Enabled (true/false, current: $(Get-ConfigValue 'Search.EnableFuzzySearch'))"
    if ($fuzzyEnabled -in @("true", "false")) {
        Set-ConfigValue -Path "Search.EnableFuzzySearch" -Value ([bool]::Parse($fuzzyEnabled)) -SaveImmediately
    }

    $threshold = Read-Host "Fuzzy Search Threshold (0.0-1.0, current: $(Get-ConfigValue 'Search.FuzzySearchThreshold'))"
    if ($threshold -match '^\d+\.?\d*$') {
        $thresholdValue = [double]$threshold
        if ($thresholdValue -ge 0.0 -and $thresholdValue -le 1.0) {
            Set-ConfigValue -Path "Search.FuzzySearchThreshold" -Value $thresholdValue -SaveImmediately
        }
    }

    Write-Host "Search settings updated!" -ForegroundColor $Global:ColorScheme.Success
    Start-Sleep -Seconds 2
}

<#
.SYNOPSIS
    Edits rule engine settings
#>
function Edit-RuleEngineSettings {
    Clear-Host
    Write-Host "`n⚙️  Rule Engine Settings" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    $enabled = Read-Host "Rule Engine Enabled (true/false, current: $(Get-ConfigValue 'Rules.Enabled'))"
    if ($enabled -in @("true", "false")) {
        Set-ConfigValue -Path "Rules.Enabled" -Value ([bool]::Parse($enabled)) -SaveImmediately
    }

    $autoExecute = Read-Host "Auto Execute Rules (true/false, current: $(Get-ConfigValue 'Rules.AutoExecuteRules'))"
    if ($autoExecute -in @("true", "false")) {
        Set-ConfigValue -Path "Rules.AutoExecuteRules" -Value ([bool]::Parse($autoExecute)) -SaveImmediately
    }

    Write-Host "Rule engine settings updated!" -ForegroundColor $Global:ColorScheme.Success
    Start-Sleep -Seconds 2
}

# Export functions
Export-ModuleMember -Function Show-SettingsMenu
