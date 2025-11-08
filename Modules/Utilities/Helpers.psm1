<#
.SYNOPSIS
    Helper utility functions for MailCleanBuddy
#>

<#
.SYNOPSIS
    Safely parses a DateTime value with culture-invariant handling
.PARAMETER DateTimeValue
    The DateTime value to parse (can be DateTime object or string)
.PARAMETER DefaultValue
    Default value to return if parsing fails (defaults to $null)
.OUTPUTS
    DateTime object or default value
#>
function ConvertTo-SafeDateTime {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        $DateTimeValue,

        [Parameter(Mandatory = $false)]
        [DateTime]$DefaultValue = [DateTime]::MinValue
    )

    if ($null -eq $DateTimeValue) {
        return $DefaultValue
    }

    try {
        # If already a DateTime object, return it
        if ($DateTimeValue -is [DateTime]) {
            return $DateTimeValue
        }

        # Try parsing as string with InvariantCulture
        if ($DateTimeValue -is [string]) {
            return [DateTime]::Parse($DateTimeValue, [System.Globalization.CultureInfo]::InvariantCulture)
        }

        # Try Get-Date cmdlet as fallback
        return (Get-Date $DateTimeValue -ErrorAction Stop)
    }
    catch {
        Write-Verbose "Failed to parse DateTime value: $DateTimeValue - Error: $($_.Exception.Message)"
        return $DefaultValue
    }
}

<#
.SYNOPSIS
    Converts HTML content to plain text
.PARAMETER HtmlContent
    The HTML content to convert
#>
function Convert-HtmlToPlainText {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$HtmlContent
    )

    if ([string]::IsNullOrWhiteSpace($HtmlContent)) {
        return ""
    }

    try {
        $plainText = $HtmlContent

        # Remove script and style tags
        $plainText = $plainText -replace '<script[^>]*>[\s\S]*?</script>', ''
        $plainText = $plainText -replace '<style[^>]*>[\s\S]*?</style>', ''

        # Convert common HTML entities
        $plainText = $plainText -replace '&nbsp;', ' '
        $plainText = $plainText -replace '&lt;', '<'
        $plainText = $plainText -replace '&gt;', '>'
        $plainText = $plainText -replace '&amp;', '&'
        $plainText = $plainText -replace '&quot;', '"'
        $plainText = $plainText -replace '&#39;', "'"
        $plainText = $plainText -replace '<br\s*/?>', "`n"
        $plainText = $plainText -replace '</p>', "`n`n"
        $plainText = $plainText -replace '</div>', "`n"

        # Remove all remaining HTML tags
        $plainText = $plainText -replace '<[^>]+>', ''

        # Clean up whitespace
        $plainText = $plainText -replace '\r\n', "`n"
        $plainText = $plainText -replace '\r', "`n"
        $plainText = $plainText -replace '\n{3,}', "`n`n"
        $plainText = $plainText.Trim()

        return $plainText
    } catch {
        Write-Warning "Error converting HTML to plain text: $($_.Exception.Message)"
        return $HtmlContent
    }
}

<#
.SYNOPSIS
    Gets a Yes/No confirmation from user
.PARAMETER Prompt
    The confirmation prompt
#>
function Get-UserConfirmation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Prompt
    )

    $confirmation = Read-Host $Prompt
    return ($confirmation -match '^(y|yes|j|ja)$')
}

<#
.SYNOPSIS
    Ensures a directory path exists
.PARAMETER Path
    The directory path
#>
function Ensure-DirectoryExists {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Verbose "Created directory: $Path"
        }
        return $true
    } catch {
        Write-Error "Could not create directory '$Path': $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Sanitizes a string for use in filenames
.PARAMETER Text
    The text to sanitize
.PARAMETER MaxLength
    Maximum length (default: 50)
#>
function Get-SafeFilename {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $false)]
        [int]$MaxLength = 50
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return "unnamed"
    }

    # Get invalid filename characters
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() + @(':', '/', '\', '?', '*', '"', '<', '>', '|')
    $regexPattern = "[{0}]" -f ([regex]::Escape(-join $invalidChars))

    # Replace invalid characters with underscore
    $safeName = $Text -replace $regexPattern, '_'

    # Trim to max length
    if ($safeName.Length -gt $MaxLength) {
        $safeName = $safeName.Substring(0, $MaxLength)
    }

    # Trim whitespace
    $safeName = $safeName.Trim()

    return $safeName
}

<#
.SYNOPSIS
    Formats file size in human-readable format
.PARAMETER Bytes
    Size in bytes
#>
function Format-FileSize {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [long]$Bytes
    )

    if ($Bytes -ge 1GB) {
        return "{0:N2} GB" -f ($Bytes / 1GB)
    } elseif ($Bytes -ge 1MB) {
        return "{0:N2} MB" -f ($Bytes / 1MB)
    } elseif ($Bytes -ge 1KB) {
        return "{0:N2} KB" -f ($Bytes / 1KB)
    } else {
        return "{0} bytes" -f $Bytes
    }
}

<#
.SYNOPSIS
    Gets a unique filename if file already exists
.PARAMETER FilePath
    The desired file path
#>
function Get-UniqueFilePath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        return $FilePath
    }

    $directory = [System.IO.Path]::GetDirectoryName($FilePath)
    $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath)

    $counter = 1
    do {
        $newFileName = "{0}_{1}{2}" -f $fileNameWithoutExt, $counter, $extension
        $newFilePath = Join-Path -Path $directory -ChildPath $newFileName
        $counter++
    } while (Test-Path $newFilePath)

    return $newFilePath
}

Export-ModuleMember -Function ConvertTo-SafeDateTime, Convert-HtmlToPlainText, Get-UserConfirmation, Ensure-DirectoryExists, Get-SafeFilename, Format-FileSize, Get-UniqueFilePath
