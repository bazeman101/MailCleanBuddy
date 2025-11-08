<#
.SYNOPSIS
    Fuzzy Search Engine for MailCleanBuddy
.DESCRIPTION
    Provides fuzzy string matching capabilities using Levenshtein distance algorithm
    Helps find emails even with typos or spelling variations
#>

<#
.SYNOPSIS
    Calculates Levenshtein distance between two strings
.DESCRIPTION
    Returns the minimum number of single-character edits needed to change one string into another
#>
function Get-LevenshteinDistance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$String1,

        [Parameter(Mandatory = $true)]
        [string]$String2
    )

    if ([string]::IsNullOrEmpty($String1)) {
        if ([string]::IsNullOrEmpty($String2)) { return 0 }
        return $String2.Length
    }

    if ([string]::IsNullOrEmpty($String2)) {
        return $String1.Length
    }

    $len1 = $String1.Length
    $len2 = $String2.Length

    # Create distance matrix
    $distance = New-Object 'int[,]' ($len1 + 1), ($len2 + 1)

    # Initialize first column and row
    for ($i = 0; $i -le $len1; $i++) {
        $distance[$i, 0] = $i
    }
    for ($j = 0; $j -le $len2; $j++) {
        $distance[0, $j] = $j
    }

    # Calculate distances
    for ($i = 1; $i -le $len1; $i++) {
        for ($j = 1; $j -le $len2; $j++) {
            $cost = if ($String1[$i - 1] -eq $String2[$j - 1]) { 0 } else { 1 }

            $distance[$i, $j] = [Math]::Min(
                [Math]::Min(
                    $distance[$i - 1, $j] + 1,      # deletion
                    $distance[$i, $j - 1] + 1       # insertion
                ),
                $distance[$i - 1, $j - 1] + $cost   # substitution
            )
        }
    }

    return $distance[$len1, $len2]
}

<#
.SYNOPSIS
    Calculates similarity ratio between two strings
.DESCRIPTION
    Returns a value between 0.0 (completely different) and 1.0 (identical)
#>
function Get-StringSimilarity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$String1,

        [Parameter(Mandatory = $true)]
        [string]$String2,

        [Parameter(Mandatory = $false)]
        [switch]$CaseInsensitive
    )

    if ($CaseInsensitive) {
        $String1 = $String1.ToLower()
        $String2 = $String2.ToLower()
    }

    if ($String1 -eq $String2) {
        return 1.0
    }

    $distance = Get-LevenshteinDistance -String1 $String1 -String2 $String2
    $maxLength = [Math]::Max($String1.Length, $String2.Length)

    if ($maxLength -eq 0) {
        return 1.0
    }

    $similarity = 1.0 - ($distance / $maxLength)
    return $similarity
}

<#
.SYNOPSIS
    Performs fuzzy search on a string
.DESCRIPTION
    Checks if search term matches target string within specified threshold
#>
function Test-FuzzyMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [Parameter(Mandatory = $true)]
        [string]$TargetString,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0.0, 1.0)]
        [double]$Threshold = 0.8,

        [Parameter(Mandatory = $false)]
        [switch]$CaseInsensitive = $true
    )

    if ([string]::IsNullOrWhiteSpace($SearchTerm) -or [string]::IsNullOrWhiteSpace($TargetString)) {
        return $false
    }

    # Get config threshold if available
    $configThreshold = Get-ConfigValue -Path "Search.FuzzySearchThreshold" -DefaultValue 0.8
    if ($Threshold -eq 0.8 -and $configThreshold -ne 0.8) {
        $Threshold = $configThreshold
    }

    # Exact match first (fastest)
    if ($CaseInsensitive) {
        if ($TargetString.ToLower() -like "*$($SearchTerm.ToLower())*") {
            return $true
        }
    } else {
        if ($TargetString -like "*$SearchTerm*") {
            return $true
        }
    }

    # Word-level fuzzy matching
    $searchWords = $SearchTerm -split '\s+' | Where-Object { $_.Length -gt 0 }
    $targetWords = $TargetString -split '\s+' | Where-Object { $_.Length -gt 0 }

    foreach ($searchWord in $searchWords) {
        $bestMatch = 0.0

        foreach ($targetWord in $targetWords) {
            $similarity = Get-StringSimilarity -String1 $searchWord -String2 $targetWord -CaseInsensitive:$CaseInsensitive
            if ($similarity -gt $bestMatch) {
                $bestMatch = $similarity
            }
        }

        if ($bestMatch -ge $Threshold) {
            return $true
        }
    }

    return $false
}

<#
.SYNOPSIS
    Performs fuzzy search on email messages
#>
function Invoke-FuzzyEmailSearch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0.0, 1.0)]
        [double]$Threshold = 0.8,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Subject", "Body", "From", "All")]
        [string]$SearchScope = "All"
    )

    try {
        # Get config values
        $enableFuzzy = Get-ConfigValue -Path "Search.EnableFuzzySearch" -DefaultValue $true
        if (-not $enableFuzzy) {
            Write-Host "Fuzzy search is disabled. Enable it in config." -ForegroundColor $Global:ColorScheme.Warning
            return @()
        }

        $configThreshold = Get-ConfigValue -Path "Search.FuzzySearchThreshold" -DefaultValue 0.8
        if ($Threshold -eq 0.8) { $Threshold = $configThreshold }

        $defaultScope = Get-ConfigValue -Path "Search.DefaultSearchScope" -DefaultValue "All"
        if ($SearchScope -eq "All") { $SearchScope = $defaultScope }

        Write-Host "`n🔍 Fuzzy Search" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host "Search term: '$SearchTerm' (threshold: $Threshold)" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "Searching..." -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Get cache
        $cache = Get-SenderCache
        if (-not $cache -or $cache.Count -eq 0) {
            Write-Host "No cache available. Please build cache first." -ForegroundColor $Global:ColorScheme.Warning
            return @()
        }

        $results = @()
        $totalMessages = 0
        $processed = 0

        # Count total messages
        foreach ($domain in $cache.Keys) {
            $totalMessages += $cache[$domain].Messages.Count
        }

        # Search through cache
        foreach ($domain in $cache.Keys) {
            foreach ($message in $cache[$domain].Messages) {
                $processed++

                if ($processed % 50 -eq 0) {
                    $percent = [Math]::Round(($processed / $totalMessages) * 100)
                    Write-Progress -Activity "Fuzzy searching emails..." `
                                   -Status "Processed $processed of $totalMessages" `
                                   -PercentComplete $percent
                }

                $isMatch = $false

                # Check Subject
                if (($SearchScope -eq "All" -or $SearchScope -eq "Subject") -and $message.Subject) {
                    if (Test-FuzzyMatch -SearchTerm $SearchTerm -TargetString $message.Subject -Threshold $Threshold) {
                        $isMatch = $true
                    }
                }

                # Check Body Preview
                if (-not $isMatch -and ($SearchScope -eq "All" -or $SearchScope -eq "Body") -and $message.BodyPreview) {
                    if (Test-FuzzyMatch -SearchTerm $SearchTerm -TargetString $message.BodyPreview -Threshold $Threshold) {
                        $isMatch = $true
                    }
                }

                # Check Sender
                if (-not $isMatch -and ($SearchScope -eq "All" -or $SearchScope -eq "From")) {
                    if ($message.SenderEmailAddress) {
                        if (Test-FuzzyMatch -SearchTerm $SearchTerm -TargetString $message.SenderEmailAddress -Threshold $Threshold) {
                            $isMatch = $true
                        }
                    }
                    if (-not $isMatch -and $message.SenderName) {
                        if (Test-FuzzyMatch -SearchTerm $SearchTerm -TargetString $message.SenderName -Threshold $Threshold) {
                            $isMatch = $true
                        }
                    }
                }

                if ($isMatch) {
                    # Prepare message for display
                    $msgId = if ($message.MessageId) { $message.MessageId } elseif ($message.Id) { $message.Id } else { $null }
                    if ($msgId) {
                        $results += [PSCustomObject]@{
                            Id                 = $msgId
                            MessageId          = $msgId
                            Subject            = if ($message.Subject) { $message.Subject } else { "(No Subject)" }
                            SenderName         = if ($message.SenderName) { $message.SenderName } else { "N/A" }
                            SenderEmailAddress = if ($message.SenderEmailAddress) { $message.SenderEmailAddress } else { "N/A" }
                            ReceivedDateTime   = $message.ReceivedDateTime
                            Size               = if ($message.Size) { $message.Size } else { 0 }
                            HasAttachments     = if ($message.HasAttachments) { $message.HasAttachments } else { $false }
                            BodyPreview        = if ($message.BodyPreview) { $message.BodyPreview } else { "" }
                        }
                    }
                }
            }
        }

        Write-Progress -Activity "Fuzzy searching emails..." -Completed

        # Save to search history
        Save-SearchToHistory -UserEmail $UserEmail -SearchTerm $SearchTerm -ResultCount $results.Count -SearchType "Fuzzy"

        Write-Host "✅ Found $($results.Count) matching email(s)" -ForegroundColor $Global:ColorScheme.Success
        Write-Host ""

        return $results
    } catch {
        Write-Error "Error in fuzzy search: $($_.Exception.Message)"
        return @()
    }
}

<#
.SYNOPSIS
    Shows fuzzy search UI
#>
function Show-FuzzySearchUI {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Clear-Host
        Write-Host "`n🔍 Fuzzy Email Search" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host "Fuzzy search helps you find emails even with typos or spelling variations." -ForegroundColor $Global:ColorScheme.Info
        Write-Host "Example: searching for 'meetng' will find 'meeting', 'meetings', etc." -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Check if enabled
        $enableFuzzy = Get-ConfigValue -Path "Search.EnableFuzzySearch" -DefaultValue $true
        if (-not $enableFuzzy) {
            Write-Host "⚠️  Fuzzy search is currently disabled in config." -ForegroundColor $Global:ColorScheme.Warning
            Write-Host "To enable: Set 'Search.EnableFuzzySearch' to true in config.json" -ForegroundColor $Global:ColorScheme.Info
            Write-Host ""
            Read-Host "Press Enter to continue"
            return
        }

        $threshold = Get-ConfigValue -Path "Search.FuzzySearchThreshold" -DefaultValue 0.8
        $searchScope = Get-ConfigValue -Path "Search.DefaultSearchScope" -DefaultValue "All"

        Write-Host "⚙️  Current Settings:" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  Threshold: $threshold (0.0 = loose, 1.0 = exact)" -ForegroundColor $Global:ColorScheme.Muted
        Write-Host "  Search scope: $searchScope" -ForegroundColor $Global:ColorScheme.Muted
        Write-Host ""

        $searchTerm = Read-Host "Enter search term (or Q to quit)"

        if ([string]::IsNullOrWhiteSpace($searchTerm) -or $searchTerm -match '^(q|quit)$') {
            return
        }

        # Perform fuzzy search
        $results = Invoke-FuzzyEmailSearch -UserEmail $UserEmail -SearchTerm $searchTerm -Threshold $threshold -SearchScope $searchScope

        if ($results.Count -eq 0) {
            Write-Host "No matching emails found." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        # Show results
        Start-Sleep -Seconds 1
        Show-StandardizedEmailListView -UserEmail $UserEmail `
                                       -Messages $results `
                                       -Title "Fuzzy Search Results: '$searchTerm' ($($results.Count) found)" `
                                       -AllowActions $true `
                                       -ViewName "FuzzySearchResults"
    } catch {
        Write-Error "Error in fuzzy search UI: $($_.Exception.Message)"
        Read-Host "Press Enter to continue"
    }
}

<#
.SYNOPSIS
    Saves search to history
#>
function Save-SearchToHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [Parameter(Mandatory = $true)]
        [int]$ResultCount,

        [Parameter(Mandatory = $false)]
        [string]$SearchType = "Standard"
    )

    try {
        $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
        $homeDir = if ($IsWindows -or $null -eq $IsWindows) { $env:USERPROFILE } else { $env:HOME }
        $historyDir = Join-Path $homeDir ".mailcleanbuddy"

        if (-not (Test-Path $historyDir)) {
            New-Item -Path $historyDir -ItemType Directory -Force | Out-Null
        }

        $historyPath = Join-Path $historyDir "search_history_$sanitizedEmail.json"

        $history = @()
        if (Test-Path $historyPath) {
            $history = Get-Content -Path $historyPath -Raw | ConvertFrom-Json
        }

        $entry = [PSCustomObject]@{
            SearchTerm = $SearchTerm
            ResultCount = $ResultCount
            SearchType = $SearchType
            Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }

        $history += $entry

        # Keep only last 20 searches
        $maxHistory = Get-ConfigValue -Path "Search.SearchHistorySize" -DefaultValue 20
        if ($history.Count -gt $maxHistory) {
            $history = $history | Select-Object -Last $maxHistory
        }

        $history | ConvertTo-Json -Depth 10 | Set-Content -Path $historyPath -Encoding UTF8
    } catch {
        Write-Verbose "Could not save search history: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Gets search history
#>
function Get-SearchHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
        $homeDir = if ($IsWindows -or $null -eq $IsWindows) { $env:USERPROFILE } else { $env:HOME }
        $historyPath = Join-Path $homeDir ".mailcleanbuddy" "search_history_$sanitizedEmail.json"

        if (-not (Test-Path $historyPath)) {
            return @()
        }

        $history = Get-Content -Path $historyPath -Raw | ConvertFrom-Json
        return $history | Sort-Object { [datetime]$_.Timestamp } -Descending
    } catch {
        Write-Warning "Could not load search history: $($_.Exception.Message)"
        return @()
    }
}

<#
.SYNOPSIS
    Shows search history
#>
function Show-SearchHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Clear-Host
        Write-Host "`n📜 Search History" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        $history = Get-SearchHistory -UserEmail $UserEmail

        if ($history.Count -eq 0) {
            Write-Host "No search history found." -ForegroundColor $Global:ColorScheme.Muted
            Read-Host "Press Enter to continue"
            return
        }

        $index = 1
        foreach ($entry in $history) {
            Write-Host "  $index. " -NoNewline -ForegroundColor $Global:ColorScheme.Muted
            Write-Host "[$($entry.Timestamp)] " -NoNewline -ForegroundColor $Global:ColorScheme.Muted
            Write-Host "$($entry.SearchTerm) " -NoNewline -ForegroundColor $Global:ColorScheme.Value
            Write-Host "($($entry.SearchType)) " -NoNewline -ForegroundColor $Global:ColorScheme.Info
            Write-Host "→ $($entry.ResultCount) results" -ForegroundColor $Global:ColorScheme.Success
            $index++
        }

        Write-Host ""
        Read-Host "Press Enter to continue"
    } catch {
        Write-Error "Error showing search history: $($_.Exception.Message)"
        Read-Host "Press Enter to continue"
    }
}

# Export functions
Export-ModuleMember -Function Get-LevenshteinDistance, Get-StringSimilarity, Test-FuzzyMatch, `
                              Invoke-FuzzyEmailSearch, Show-FuzzySearchUI, Save-SearchToHistory, `
                              Get-SearchHistory, Show-SearchHistory
