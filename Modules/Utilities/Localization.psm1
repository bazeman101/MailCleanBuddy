<#
.SYNOPSIS
    Localization module for MailCleanBuddy
.DESCRIPTION
    Handles loading and retrieving localized strings from localizations.json
#>

# Script-level variables
$Script:LocalizedStrings = $null
$Script:SelectedLanguage = "nl"
$Script:LocalizationFilePath = $null

<#
.SYNOPSIS
    Loads localization strings from JSON file
.PARAMETER SelectedLang
    The language code to load (nl, en, de, fr)
.PARAMETER FilePath
    Path to the localizations.json file
#>
function Initialize-Localization {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$SelectedLang = "nl",

        [Parameter(Mandatory = $false)]
        [string]$FilePath
    )

    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        $Script:LocalizationFilePath = Join-Path -Path $PSScriptRoot -ChildPath "..\..\localizations.json"
    } else {
        $Script:LocalizationFilePath = $FilePath
    }

    if (-not (Test-Path $Script:LocalizationFilePath)) {
        Write-Warning "Localization file not found at: $Script:LocalizationFilePath. Using fallback internal strings (limited)."
        # Fallback naar minimale set
        $Script:LocalizedStrings = @{
            "nl" = @{ "mainMenu_title" = "MailCleanBuddy - Hoofdmenu voor {0}"; "mainMenu_optionQ" = "Q. Afsluiten" }
            "en" = @{ "mainMenu_title" = "MailCleanBuddy - Main Menu for {0}"; "mainMenu_optionQ" = "Q. Quit" }
        }
        if ($Script:LocalizedStrings.ContainsKey($SelectedLang)) {
            $Script:LocalizedStrings = $Script:LocalizedStrings[$SelectedLang]
        } else {
            $Script:LocalizedStrings = $Script:LocalizedStrings["nl"]
        }
        $Script:SelectedLanguage = $SelectedLang
        return
    }

    try {
        $jsonContent = Get-Content -Path $Script:LocalizationFilePath -Raw -ErrorAction Stop
        $allLocalizations = ConvertFrom-Json -InputObject $jsonContent -ErrorAction Stop

        if ($allLocalizations.PSObject.Properties.Name -contains $SelectedLang) {
            $Script:LocalizedStrings = $allLocalizations.$SelectedLang
            $Script:SelectedLanguage = $SelectedLang
        } elseif ($allLocalizations.PSObject.Properties.Name -contains "nl") {
            Write-Warning "Language '$SelectedLang' not found in localization file. Falling back to Dutch (nl)."
            $Script:LocalizedStrings = $allLocalizations.nl
            $Script:SelectedLanguage = "nl"
        } else {
            Write-Error "Default language 'nl' not found in localization file. Cannot load UI strings."
            throw "Critical localization error."
        }
    } catch {
        Write-Error "Error loading or parsing localization file '$Script:LocalizationFilePath': $($_.Exception.Message)"
        throw "Critical localization error."
    }
}

<#
.SYNOPSIS
    Retrieves a localized string by key
.PARAMETER Key
    The localization key
.PARAMETER FormatArgs
    Optional format arguments for string formatting
#>
function Get-LocalizedString {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $false)]
        [object[]]$FormatArgs = @()
    )

    if ($Script:LocalizedStrings) {
        $keyExists = $false
        $localizedString = $null

        if ($Script:LocalizedStrings -is [hashtable]) {
            if ($Script:LocalizedStrings.ContainsKey($Key)) {
                $keyExists = $true
                $localizedString = $Script:LocalizedStrings[$Key]
            }
        } elseif ($Script:LocalizedStrings -is [System.Management.Automation.PSCustomObject]) {
            if ($Script:LocalizedStrings.PSObject.Properties.Name -contains $Key) {
                $keyExists = $true
                $localizedString = $Script:LocalizedStrings.$Key
            }
        }

        if ($keyExists) {
            if ($FormatArgs.Count -gt 0) {
                try {
                    return ($localizedString -f $FormatArgs)
                } catch {
                    Write-Warning "Error formatting localized string for key '$Key' with args '$($FormatArgs -join ', ')'. Raw string returned. Error: $($_.Exception.Message)"
                    return $localizedString
                }
            } else {
                return $localizedString
            }
        }
    }

    Write-Warning "Localization key '$Key' not found for language '$($Script:SelectedLanguage)'. Returning key itself."
    return $Key
}

<#
.SYNOPSIS
    Gets the current selected language
#>
function Get-CurrentLanguage {
    return $Script:SelectedLanguage
}

# Export functions
Export-ModuleMember -Function Initialize-Localization, Get-LocalizedString, Get-CurrentLanguage
