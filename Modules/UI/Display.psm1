<#
.SYNOPSIS
    Display helper module for MailCleanBuddy
.DESCRIPTION
    Provides display utilities for menus and lists
#>

# Import required modules

<#
.SYNOPSIS
    Sets console window size
.PARAMETER Width
    Window width
.PARAMETER Height
    Window height
#>
function Set-ConsoleSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$Width = 150,

        [Parameter(Mandatory = $false)]
        [int]$Height = 55
    )

    try {
        if ($Host.UI.GetType().Name -notmatch "ConsoleHostUserInterface") {
            Write-Verbose "Not an interactive console, skipping size adjustment"
            return
        }

        $currentWindowSize = $Host.UI.RawUI.WindowSize
        $bufferSize = $Host.UI.RawUI.BufferSize

        # Adjust buffer size if needed
        if ($bufferSize.Width -lt $Width) {
            $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size ($Width, $bufferSize.Height)
        }

        $newBufferHeight = [Math]::Max($bufferSize.Height, $Height)
        if ($Host.UI.RawUI.BufferSize.Width -lt $Width -or $Host.UI.RawUI.BufferSize.Height -lt $newBufferHeight) {
            $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size (
                [Math]::Max($Host.UI.RawUI.BufferSize.Width, $Width),
                $newBufferHeight
            )
        }

        # Set window size
        $Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size ($Width, $Height)

        Write-Verbose "Console size set to Width: $Width, Height: $Height"

    } catch {
        Write-Warning "Could not set console window size: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Displays a formatted header
.PARAMETER Title
    Header title
.PARAMETER Width
    Header width
#>
function Show-Header {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [int]$Width = 80
    )

    Initialize-ColorScheme
    $colors = Get-ColorScheme

    if ($null -ne $colors) {
        $Host.UI.RawUI.ForegroundColor = $colors.ForegroundColor
        $Host.UI.RawUI.BackgroundColor = $colors.BackgroundColor
    }

    $separator = "=" * $Width
    Write-Host $separator
    Write-Host $Title.PadRight($Width)
    Write-Host $separator
}

<#
.SYNOPSIS
    Displays a formatted list
.PARAMETER Items
    Array of items to display
.PARAMETER SelectedIndex
    Index of selected item
.PARAMETER StartIndex
    Start index for pagination
.PARAMETER PageSize
    Number of items per page
#>
function Show-List {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Items,

        [Parameter(Mandatory = $false)]
        [int]$SelectedIndex = 0,

        [Parameter(Mandatory = $false)]
        [int]$StartIndex = 0,

        [Parameter(Mandatory = $false)]
        [int]$PageSize = 30
    )

    $colors = Get-ColorScheme

    $endIndex = [Math]::Min($StartIndex + $PageSize, $Items.Count)

    for ($i = $StartIndex; $i -lt $endIndex; $i++) {
        $item = $Items[$i]
        $displayText = if ($item -is [string]) { $item } else { $item.ToString() }

        if ($i -eq $SelectedIndex) {
            Write-Host "> $displayText" -ForegroundColor $colors.SelectedForegroundColor -BackgroundColor $colors.SelectedBackgroundColor
        } else {
            Write-Host "  $displayText" -ForegroundColor $colors.ForegroundColor
        }
    }
}

<#
.SYNOPSIS
    Displays instructions
.PARAMETER Instructions
    Array of instruction strings
#>
function Show-Instructions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Instructions
    )

    $colors = Get-ColorScheme
    Write-Host ""
    foreach ($instruction in $Instructions) {
        Write-Host $instruction -ForegroundColor $colors.InstructionColor
    }
}

<#
.SYNOPSIS
    Displays a warning message
.PARAMETER Message
    Warning message
#>
function Show-WarningMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $colors = Get-ColorScheme
    Write-Host $Message -ForegroundColor $colors.WarningColor
}

<#
.SYNOPSIS
    Displays an error message
.PARAMETER Message
    Error message
#>
function Show-ErrorMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $colors = Get-ColorScheme
    Write-Host $Message -ForegroundColor $colors.ErrorColor
}

<#
.SYNOPSIS
    Displays a success message
.PARAMETER Message
    Success message
#>
function Show-SuccessMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $colors = Get-ColorScheme
    Write-Host $Message -ForegroundColor $colors.SuccessColor
}

<#
.SYNOPSIS
    Reads a key press from user
.PARAMETER IncludeKeyDown
    Include key down events
#>
function Read-KeyPress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$IncludeKeyDown
    )

    $options = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho
    if ($IncludeKeyDown) {
        $options = $options -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
    }

    return $Host.UI.RawUI.ReadKey($options)
}

<#
.SYNOPSIS
    Waits for Enter key press
.PARAMETER Message
    Optional message to display
#>
function Wait-EnterKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Message = "Press Enter to continue..."
    )

    $colors = Get-ColorScheme
    Write-Host $Message -ForegroundColor $colors.InstructionColor
    Read-Host | Out-Null
}

Export-ModuleMember -Function Set-ConsoleSize, Show-Header, Show-List, Show-Instructions, Show-WarningMessage, `
                              Show-ErrorMessage, Show-SuccessMessage, Read-KeyPress, Wait-EnterKey
