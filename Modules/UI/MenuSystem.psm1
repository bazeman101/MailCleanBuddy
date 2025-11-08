<#
.SYNOPSIS
    Menu system module for MailCleanBuddy
.DESCRIPTION
    Provides interactive menu functionality
#>

# Import required modules

<#
.SYNOPSIS
    Shows a simple menu with options
.PARAMETER Title
    Menu title
.PARAMETER Options
    Array of menu options (hashtable with 'Key' and 'Description')
.PARAMETER AllowEscape
    Allow Escape key to exit
#>
function Show-SimpleMenu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [hashtable[]]$Options,

        [Parameter(Mandatory = $false)]
        [switch]$AllowEscape
    )

    $colors = Get-ColorScheme
    Set-DefaultColors

    Clear-Host
    Show-Header -Title $Title -Width 80

    Write-Host ""
    foreach ($option in $Options) {
        Write-Host "  $($option.Key). $($option.Description)" -ForegroundColor $colors.ForegroundColor
    }
    Write-Host ""

    if ($AllowEscape) {
        Write-Host "  ESC or Q. Quit/Back" -ForegroundColor $colors.InstructionColor
    }

    Write-Host ""
    $selection = Read-Host "Select an option"

    return $selection
}

<#
.SYNOPSIS
    Shows a selectable list with arrow key navigation
.PARAMETER Title
    List title
.PARAMETER Items
    Array of items (can be strings or objects)
.PARAMETER DisplayProperty
    Property name to display if items are objects
.PARAMETER PageSize
    Number of items to display per page
.PARAMETER AllowMultiSelect
    Allow selecting multiple items with spacebar
#>
function Show-SelectableList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [array]$Items,

        [Parameter(Mandatory = $false)]
        [string]$DisplayProperty,

        [Parameter(Mandatory = $false)]
        [int]$PageSize = 30,

        [Parameter(Mandatory = $false)]
        [switch]$AllowMultiSelect
    )

    if ($Items.Count -eq 0) {
        Show-WarningMessage "No items to display"
        Wait-EnterKey
        return $null
    }

    $colors = Get-ColorScheme
    $selectedIndex = 0
    $topDisplayIndex = 0
    $selectedItems = @()
    $exitLoop = $false
    $returnValue = $null

    while (-not $exitLoop) {
        Set-DefaultColors
        Clear-Host
        Show-Header -Title $Title -Width 80

        # Display items
        $endIndex = [Math]::Min($topDisplayIndex + $PageSize, $Items.Count)

        for ($i = $topDisplayIndex; $i -lt $endIndex; $i++) {
            $item = $Items[$i]
            $displayText = if ($DisplayProperty) {
                $item.$DisplayProperty
            } elseif ($item -is [string]) {
                $item
            } else {
                $item.ToString()
            }

            # Add selection marker if multi-select
            $prefix = "  "
            if ($AllowMultiSelect -and $selectedItems -contains $i) {
                $prefix = "[*] "
            }

            if ($i -eq $selectedIndex) {
                Write-Host "$prefix$displayText" -ForegroundColor $colors.SelectedForegroundColor -BackgroundColor $colors.SelectedBackgroundColor
            } else {
                Write-Host "$prefix$displayText" -ForegroundColor $colors.ForegroundColor
            }
        }

        # Show pagination info
        Write-Host ""
        Write-Host "Showing items $($topDisplayIndex + 1) - $endIndex of $($Items.Count)" -ForegroundColor $colors.InstructionColor

        # Show instructions
        $instructions = @(
            "Use Up/Down arrows to navigate, Enter to select"
        )
        if ($AllowMultiSelect) {
            $instructions += "Spacebar to toggle selection, A to select all, N to deselect all"
        }
        $instructions += "ESC or Q to cancel"

        Show-Instructions -Instructions $instructions

        if ($AllowMultiSelect -and $selectedItems.Count -gt 0) {
            Write-Host ""
            Write-Host "Selected: $($selectedItems.Count) items" -ForegroundColor $colors.HighlightColor
        }

        # Read key
        $key = Read-KeyPress -IncludeKeyDown

        switch ($key.VirtualKeyCode) {
            38 { # Up arrow
                if ($selectedIndex -gt 0) {
                    $selectedIndex--
                    if ($selectedIndex -lt $topDisplayIndex) {
                        $topDisplayIndex = $selectedIndex
                    }
                }
            }
            40 { # Down arrow
                if ($selectedIndex -lt ($Items.Count - 1)) {
                    $selectedIndex++
                    if ($selectedIndex -ge ($topDisplayIndex + $PageSize)) {
                        $topDisplayIndex = $selectedIndex - $PageSize + 1
                    }
                }
            }
            13 { # Enter
                if ($AllowMultiSelect) {
                    if ($selectedItems.Count -gt 0) {
                        $returnValue = $selectedItems | ForEach-Object { $Items[$_] }
                    } else {
                        $returnValue = @($Items[$selectedIndex])
                    }
                } else {
                    $returnValue = $Items[$selectedIndex]
                }
                $exitLoop = $true
            }
            27 { # Escape
                $returnValue = $null
                $exitLoop = $true
            }
            32 { # Spacebar
                if ($AllowMultiSelect) {
                    if ($selectedItems -contains $selectedIndex) {
                        $selectedItems = $selectedItems | Where-Object { $_ -ne $selectedIndex }
                    } else {
                        $selectedItems += $selectedIndex
                    }
                }
            }
            default {
                $char = $key.Character.ToString().ToUpper()
                if ($char -eq 'Q') {
                    $returnValue = $null
                    $exitLoop = $true
                } elseif ($AllowMultiSelect -and $char -eq 'A') {
                    # Select all
                    $selectedItems = 0..($Items.Count - 1)
                } elseif ($AllowMultiSelect -and $char -eq 'N') {
                    # Deselect all
                    $selectedItems = @()
                }
            }
        }
    }

    return $returnValue
}

<#
.SYNOPSIS
    Shows a time filter selection menu
#>
function Show-TimeFilterMenu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Title = "Select Time Filter"
    )

    $options = @(
        @{ Key = "1"; Description = "Last Day"; Value = "LastDay" }
        @{ Key = "2"; Description = "Last 7 Days"; Value = "Last7Days" }
        @{ Key = "3"; Description = "Last 30 Days"; Value = "Last30Days" }
        @{ Key = "4"; Description = "Last 90 Days"; Value = "Last90Days" }
        @{ Key = "5"; Description = "Last Week"; Value = "LastWeek" }
        @{ Key = "6"; Description = "Last Month"; Value = "LastMonth" }
        @{ Key = "7"; Description = "All"; Value = "All" }
    )

    $selection = Show-SimpleMenu -Title $Title -Options $options -AllowEscape

    $selectedOption = $options | Where-Object { $_.Key -eq $selection } | Select-Object -First 1

    if ($selectedOption) {
        return $selectedOption.Value
    }

    return $null
}

<#
.SYNOPSIS
    Shows a confirmation dialog
.PARAMETER Message
    Confirmation message
.PARAMETER DefaultYes
    Default to Yes
#>
function Show-Confirmation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [switch]$DefaultYes
    )

    $colors = Get-ColorScheme

    Write-Host ""
    Write-Host $Message -ForegroundColor $colors.WarningColor

    $prompt = if ($DefaultYes) { "[Y/n]" } else { "[y/N]" }
    $response = Read-Host $prompt

    if ([string]::IsNullOrWhiteSpace($response)) {
        return $DefaultYes.IsPresent
    }

    return $response -match '^(y|yes|j|ja)$'
}

Export-ModuleMember -Function Show-SimpleMenu, Show-SelectableList, Show-TimeFilterMenu, Show-Confirmation
