<#
.SYNOPSIS
    Color scheme definitions for MailCleanBuddy UI
.DESCRIPTION
    Provides CGA-style color scheme (green on black) for consistent UI
#>

# CGA Color Scheme
$Script:ColorScheme = @{
    BackgroundColor           = [System.ConsoleColor]::Black
    ForegroundColor           = [System.ConsoleColor]::Green
    SelectedBackgroundColor   = [System.ConsoleColor]::Green
    SelectedForegroundColor   = [System.ConsoleColor]::Black
    InstructionColor          = [System.ConsoleColor]::White
    WarningColor              = [System.ConsoleColor]::Red
    HighlightColor            = [System.ConsoleColor]::Yellow
    ErrorColor                = [System.ConsoleColor]::Red
    SuccessColor              = [System.ConsoleColor]::Green
}

# Initialize Global ColorScheme for module-wide access
$Global:ColorScheme = @{
    Border         = [System.ConsoleColor]::DarkGreen
    Error          = [System.ConsoleColor]::Red
    Header         = [System.ConsoleColor]::Cyan
    Highlight      = [System.ConsoleColor]::Yellow
    Info           = [System.ConsoleColor]::Cyan
    Label          = [System.ConsoleColor]::Gray
    Muted          = [System.ConsoleColor]::DarkGray
    Normal         = [System.ConsoleColor]::Green
    SectionHeader  = [System.ConsoleColor]::Yellow
    Success        = [System.ConsoleColor]::Green
    Value          = [System.ConsoleColor]::White
    Warning        = [System.ConsoleColor]::Yellow
}

<#
.SYNOPSIS
    Gets the color scheme
#>
function Get-ColorScheme {
    return $Script:ColorScheme
}

<#
.SYNOPSIS
    Ensures Global ColorScheme is initialized, returns safe color value
.PARAMETER ColorName
    Name of the color property to retrieve
.PARAMETER Fallback
    Fallback color if ColorScheme is not initialized (default: White)
#>
function Get-SafeColor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ColorName,

        [Parameter(Mandatory = $false)]
        [System.ConsoleColor]$Fallback = [System.ConsoleColor]::White
    )

    if ($null -eq $Global:ColorScheme) {
        Write-Verbose "ColorScheme not initialized, using fallback color: $Fallback"
        return $Fallback
    }

    if ($Global:ColorScheme.ContainsKey($ColorName)) {
        return $Global:ColorScheme[$ColorName]
    }

    Write-Verbose "ColorScheme does not contain key '$ColorName', using fallback: $Fallback"
    return $Fallback
}

<#
.SYNOPSIS
    Ensures Global ColorScheme is initialized with defaults
#>
function Initialize-ColorScheme {
    if ($null -eq $Global:ColorScheme) {
        Write-Verbose "Initializing Global ColorScheme with defaults"
        $Global:ColorScheme = @{
            Border         = [System.ConsoleColor]::DarkGreen
            Error          = [System.ConsoleColor]::Red
            Header         = [System.ConsoleColor]::Cyan
            Highlight      = [System.ConsoleColor]::Yellow
            Info           = [System.ConsoleColor]::Cyan
            Label          = [System.ConsoleColor]::Gray
            Muted          = [System.ConsoleColor]::DarkGray
            Normal         = [System.ConsoleColor]::Green
            SectionHeader  = [System.ConsoleColor]::Yellow
            Success        = [System.ConsoleColor]::Green
            Value          = [System.ConsoleColor]::White
            Warning        = [System.ConsoleColor]::Yellow
        }
    }
}

<#
.SYNOPSIS
    Sets console colors to default scheme
#>
function Set-DefaultColors {
    Initialize-ColorScheme
    $Host.UI.RawUI.ForegroundColor = $Script:ColorScheme.ForegroundColor
    $Host.UI.RawUI.BackgroundColor = $Script:ColorScheme.BackgroundColor
}

<#
.SYNOPSIS
    Resets console colors to original
#>
function Reset-ConsoleColors {
    $Host.UI.RawUI.ForegroundColor = [System.ConsoleColor]::Gray
    $Host.UI.RawUI.BackgroundColor = [System.ConsoleColor]::Black
}

Export-ModuleMember -Function Get-ColorScheme, Get-SafeColor, Initialize-ColorScheme, Set-DefaultColors, Reset-ConsoleColors
