<#
.SYNOPSIS
    Email Search module for MailCleanBuddy
.DESCRIPTION
    Provides email search and recent emails viewing functionality
#>

# Function: Get-SearchSuggestions
function Get-SearchSuggestions {
    <#
    .SYNOPSIS
        Gets search suggestions from cache
    .PARAMETER UserEmail
        User email address
    .OUTPUTS
        Hashtable with suggestions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    $suggestions = @{
        TopSenders = @()
        CommonSubjectKeywords = @()
        TotalEmails = 0
        EmailsWithAttachments = 0
        EmailsWithoutAttachments = 0
    }

    $cache = Get-SenderCache
    if (-not $cache -or $cache.Count -eq 0) {
        return $suggestions
    }

    # Get top senders
    $senderStats = @{}
    $subjectWords = @{}

    foreach ($domain in $cache.Keys) {
        if (-not $cache[$domain].Messages) { continue }

        $count = $cache[$domain].Messages.Count
        $suggestions.TotalEmails += $count

        # Count attachments
        $withAttachments = ($cache[$domain].Messages | Where-Object { $_.HasAttachments }).Count
        $suggestions.EmailsWithAttachments += $withAttachments
        $suggestions.EmailsWithoutAttachments += ($count - $withAttachments)

        # Track senders
        if (-not $senderStats.ContainsKey($domain)) {
            $senderStats[$domain] = 0
        }
        $senderStats[$domain] += $count

        # Extract common subject keywords
        foreach ($msg in $cache[$domain].Messages) {
            if ($msg.Subject) {
                # Split subject into words, filter common words
                $words = $msg.Subject -split '\s+' | Where-Object {
                    $_.Length -gt 3 -and
                    $_ -notmatch '^(the|and|for|with|from|this|that|your|has|been|are|was|were|bij|van|voor|een|het|de)$'
                }
                foreach ($word in $words) {
                    $cleanWord = $word -replace '[^\w]', ''
                    if ($cleanWord.Length -gt 3) {
                        if (-not $subjectWords.ContainsKey($cleanWord.ToLower())) {
                            $subjectWords[$cleanWord.ToLower()] = 0
                        }
                        $subjectWords[$cleanWord.ToLower()]++
                    }
                }
            }
        }
    }

    # Get top 10 senders
    $suggestions.TopSenders = $senderStats.GetEnumerator() |
        Sort-Object Value -Descending |
        Select-Object -First 10 |
        ForEach-Object { [PSCustomObject]@{ Sender = $_.Key; Count = $_.Value } }

    # Get top 10 subject keywords
    $suggestions.CommonSubjectKeywords = $subjectWords.GetEnumerator() |
        Sort-Object Value -Descending |
        Select-Object -First 10 |
        ForEach-Object { [PSCustomObject]@{ Keyword = $_.Key; Count = $_.Value } }

    return $suggestions
}

# Function: Invoke-EmailSearch
function Invoke-EmailSearch {
    <#
    .SYNOPSIS
        Advanced email search with suggestions and filters
    .PARAMETER UserEmail
        User email address
    .PARAMETER SearchTerm
        Search term to look for
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $false)]
        [string]$SearchTerm
    )

    try {
        Clear-Host
        Write-Host "`nAdvanced Email Search" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Get search suggestions
        Write-Host "Loading search suggestions..." -ForegroundColor $Global:ColorScheme.Info
        $suggestions = Get-SearchSuggestions -UserEmail $UserEmail

        if ($suggestions.TotalEmails -eq 0) {
            Write-Host "No cache data found. Please build the mailbox cache first." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        Clear-Host
        Write-Host "`nAdvanced Email Search" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Show overview
        Write-Host "📊 Mailbox Overview" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  Total emails: $($suggestions.TotalEmails)" -ForegroundColor $Global:ColorScheme.Normal
        Write-Host "  With attachments: $($suggestions.EmailsWithAttachments)" -ForegroundColor $Global:ColorScheme.Normal
        Write-Host "  Without attachments: $($suggestions.EmailsWithoutAttachments)" -ForegroundColor $Global:ColorScheme.Normal
        Write-Host ""

        # Show top senders
        if ($suggestions.TopSenders.Count -gt 0) {
            Write-Host "👥 Top 10 Senders" -ForegroundColor $Global:ColorScheme.SectionHeader
            $index = 1
            foreach ($sender in $suggestions.TopSenders) {
                Write-Host "  $index. " -NoNewline -ForegroundColor $Global:ColorScheme.Muted
                Write-Host "$($sender.Sender) " -NoNewline -ForegroundColor $Global:ColorScheme.Value
                Write-Host "($($sender.Count) emails)" -ForegroundColor $Global:ColorScheme.Muted
                $index++
            }
            Write-Host ""
        }

        # Show common subject keywords
        if ($suggestions.CommonSubjectKeywords.Count -gt 0) {
            Write-Host "🔤 Common Subject Keywords" -ForegroundColor $Global:ColorScheme.SectionHeader
            $keywordLine = "  "
            foreach ($kw in $suggestions.CommonSubjectKeywords) {
                $keywordLine += "$($kw.Keyword)($($kw.Count)), "
            }
            Write-Host $keywordLine.TrimEnd(', ') -ForegroundColor $Global:ColorScheme.Info
            Write-Host ""
        }

        # Quick filters
        Write-Host "⚡ Quick Filters" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  [A] Search emails WITH attachments" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  [N] Search emails WITHOUT attachments" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  [S] Search by specific sender (choose from top 10)" -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Search options
        Write-Host "🔍 Search Options" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  • Type keyword to search (supports regex: /pattern/)" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  • Use quotes for exact match: \"exact phrase\"" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  • Type 'Q' to quit" -ForegroundColor $Global:ColorScheme.Muted
        Write-Host ""

        $searchInput = Read-Host "Enter search term or quick filter (A/N/S)"

        if ([string]::IsNullOrWhiteSpace($searchInput) -or $searchInput -match '^(q|quit)$') {
            return
        }

        # Handle quick filters
        if ($searchInput.ToUpper() -eq 'A') {
            # Search with attachments
            Invoke-AttachmentSearch -UserEmail $UserEmail -HasAttachments $true
            return
        }
        elseif ($searchInput.ToUpper() -eq 'N') {
            # Search without attachments
            Invoke-AttachmentSearch -UserEmail $UserEmail -HasAttachments $false
            return
        }
        elseif ($searchInput.ToUpper() -eq 'S') {
            # Search by sender
            Write-Host ""
            Write-Host "Select sender (1-10) or type sender email:" -ForegroundColor $Global:ColorScheme.Info
            $senderInput = Read-Host

            if ($senderInput -match '^\d+$' -and [int]$senderInput -ge 1 -and [int]$senderInput -le $suggestions.TopSenders.Count) {
                $selectedSender = $suggestions.TopSenders[[int]$senderInput - 1].Sender
                Invoke-SenderSearch -UserEmail $UserEmail -Sender $selectedSender
            } else {
                Invoke-SenderSearch -UserEmail $UserEmail -Sender $senderInput
            }
            return
        }

        # Perform search
        Invoke-AdvancedSearch -UserEmail $UserEmail -SearchTerm $searchInput
    }
    catch {
        Write-Error "Error in email search: $($_.Exception.Message)"
        Write-Host "`nAn error occurred while searching emails." -ForegroundColor $Global:ColorScheme.Error
        Write-Host $_.Exception.Message -ForegroundColor $Global:ColorScheme.Error
        Read-Host "Press Enter to continue"
    }
}

# Function: Invoke-AdvancedSearch
function Invoke-AdvancedSearch {
    <#
    .SYNOPSIS
        Performs advanced search with regex support
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$SearchTerm
    )

    try {
        Clear-Host
        Write-Host "`nSearching..." -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Check if regex pattern
        $isRegex = $false
        $pattern = $SearchTerm
        if ($SearchTerm -match '^/(.+)/$') {
            $isRegex = $true
            $pattern = $matches[1]
            Write-Host "Using regex pattern: $pattern" -ForegroundColor $Global:ColorScheme.Info
        }
        # Check if exact match
        elseif ($SearchTerm -match '^"(.+)"$') {
            $pattern = $matches[1]
            Write-Host "Searching for exact phrase: $pattern" -ForegroundColor $Global:ColorScheme.Info
        }
        else {
            Write-Host "Searching for: $SearchTerm" -ForegroundColor $Global:ColorScheme.Info
        }

        Write-Host "Please wait..." -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Search using Microsoft Graph
        $messages = Search-GraphMessages -UserId $UserEmail -SearchTerm $pattern -Top 100

        if (-not $messages -or $messages.Count -eq 0) {
            Write-Host "No emails found matching: $SearchTerm" -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        # If regex, filter results
        if ($isRegex) {
            $messages = $messages | Where-Object {
                ($_.Subject -match $pattern) -or
                ($_.BodyPreview -match $pattern) -or
                ($_.From.EmailAddress.Address -match $pattern)
            }
        }

        if ($messages.Count -eq 0) {
            Write-Host "No emails found matching regex pattern: $pattern" -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        Write-Host "Found $($messages.Count) email(s)" -ForegroundColor $Global:ColorScheme.Success
        Start-Sleep -Seconds 1

        # Prepare messages for display
        $messagesForView = @()
        foreach ($msg in $messages) {
            $messagesForView += [PSCustomObject]@{
                Id                 = $msg.Id
                ReceivedDateTime   = $msg.ReceivedDateTime
                Subject            = $msg.Subject
                SenderName         = if ($msg.From -and $msg.From.EmailAddress) { $msg.From.EmailAddress.Name } else { "N/A" }
                SenderEmailAddress = if ($msg.From -and $msg.From.EmailAddress) { $msg.From.EmailAddress.Address } else { "N/A" }
                Size               = if ($msg.PSObject.Properties['Size']) { $msg.Size } else { 0 }
                HasAttachments     = if ($msg.PSObject.Properties['HasAttachments']) { $msg.HasAttachments } else { $false }
                BodyPreview        = if ($msg.PSObject.Properties['BodyPreview']) { $msg.BodyPreview } else { "" }
            }
        }

        # Display using standardized email list view
        Show-StandardizedEmailListView -UserEmail $UserEmail `
                                       -Messages $messagesForView `
                                       -Title "Search Results: $SearchTerm" `
                                       -AllowActions $true `
                                       -ViewName "SearchResults"
    }
    catch {
        Write-Error "Error performing search: $($_.Exception.Message)"
        Read-Host "Press Enter to continue"
    }
}

# Function: Invoke-AttachmentSearch
function Invoke-AttachmentSearch {
    <#
    .SYNOPSIS
        Searches emails by attachment status
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [bool]$HasAttachments
    )

    try {
        $filterText = if ($HasAttachments) { "WITH" } else { "WITHOUT" }
        Write-Host "`nSearching for emails $filterText attachments..." -ForegroundColor $Global:ColorScheme.Info

        $cache = Get-SenderCache
        $matchingMessages = @()

        foreach ($domain in $cache.Keys) {
            if (-not $cache[$domain].Messages) { continue }

            foreach ($msg in $cache[$domain].Messages) {
                if ($msg.HasAttachments -eq $HasAttachments) {
                    $msgId = if ($msg.MessageId) { $msg.MessageId } elseif ($msg.Id) { $msg.Id } else { $null }
                    if ($msgId) {
                        $matchingMessages += [PSCustomObject]@{
                            Id                 = $msgId
                            MessageId          = $msgId
                            Subject            = if ($msg.Subject) { $msg.Subject } else { "(No Subject)" }
                            SenderName         = if ($msg.SenderName) { $msg.SenderName } else { "N/A" }
                            SenderEmailAddress = if ($msg.SenderEmailAddress) { $msg.SenderEmailAddress } else { "N/A" }
                            ReceivedDateTime   = $msg.ReceivedDateTime
                            Size               = if ($msg.Size) { $msg.Size } else { 0 }
                            HasAttachments     = $msg.HasAttachments
                            BodyPreview        = if ($msg.BodyPreview) { $msg.BodyPreview } else { "" }
                        }
                    }
                }
            }
        }

        if ($matchingMessages.Count -eq 0) {
            Write-Host "No emails found $filterText attachments." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        Write-Host "Found $($matchingMessages.Count) email(s) $filterText attachments" -ForegroundColor $Global:ColorScheme.Success
        Start-Sleep -Seconds 1

        # Sort by date
        $matchingMessages = $matchingMessages | Sort-Object ReceivedDateTime -Descending

        Show-StandardizedEmailListView -UserEmail $UserEmail `
                                       -Messages $matchingMessages `
                                       -Title "Emails $filterText Attachments ($($matchingMessages.Count) found)" `
                                       -AllowActions $true `
                                       -ViewName "AttachmentSearch"
    }
    catch {
        Write-Error "Error searching by attachment: $($_.Exception.Message)"
        Read-Host "Press Enter to continue"
    }
}

# Function: Invoke-SenderSearch
function Invoke-SenderSearch {
    <#
    .SYNOPSIS
        Searches emails by sender
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$Sender
    )

    try {
        Write-Host "`nSearching for emails from: $Sender..." -ForegroundColor $Global:ColorScheme.Info

        $cache = Get-SenderCache
        $senderKey = $Sender.ToLower()

        # Try to find sender in cache
        $senderData = $null
        foreach ($domain in $cache.Keys) {
            if ($domain.ToLower() -eq $senderKey -or $domain.ToLower() -like "*$senderKey*") {
                $senderData = $cache[$domain]
                break
            }
        }

        if (-not $senderData -or -not $senderData.Messages) {
            Write-Host "No emails found from: $Sender" -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        # Prepare messages
        $messagesForView = @()
        foreach ($msg in $senderData.Messages) {
            $msgId = if ($msg.MessageId) { $msg.MessageId } elseif ($msg.Id) { $msg.Id } else { $null }
            if ($msgId) {
                $messagesForView += [PSCustomObject]@{
                    Id                 = $msgId
                    MessageId          = $msgId
                    ReceivedDateTime   = $msg.ReceivedDateTime
                    Subject            = if ($msg.Subject) { $msg.Subject } else { "(No Subject)" }
                    SenderName         = if ($msg.SenderName) { $msg.SenderName } else { "N/A" }
                    SenderEmailAddress = if ($msg.SenderEmailAddress) { $msg.SenderEmailAddress } else { "N/A" }
                    Size               = if ($msg.Size) { $msg.Size } else { 0 }
                    HasAttachments     = if ($msg.HasAttachments) { $msg.HasAttachments } else { $false }
                    BodyPreview        = if ($msg.BodyPreview) { $msg.BodyPreview } else { "" }
                }
            }
        }

        Write-Host "Found $($messagesForView.Count) email(s) from $Sender" -ForegroundColor $Global:ColorScheme.Success
        Start-Sleep -Seconds 1

        Show-StandardizedEmailListView -UserEmail $UserEmail `
                                       -Messages $messagesForView `
                                       -Title "Emails from: $Sender ($($messagesForView.Count) emails)" `
                                       -AllowActions $true `
                                       -ViewName "SenderSearch_$Sender"
    }
    catch {
        Write-Error "Error searching by sender: $($_.Exception.Message)"
        Read-Host "Press Enter to continue"
    }
}

# Function: Show-RecentEmails
function Show-RecentEmails {
    <#
    .SYNOPSIS
        Shows the most recent emails
    .PARAMETER UserEmail
        User email address
    .PARAMETER Count
        Number of recent emails to show (default: 100)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $false)]
        [int]$Count = 100
    )

    try {
        Clear-Host
        Write-Host "`nRecent Emails" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host "Fetching $Count most recent emails..." -ForegroundColor $Global:ColorScheme.Info
        Write-Host "Please wait..." -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Get recent messages using Graph API
        $messages = Get-GraphMessages -UserId $UserEmail -Top $Count -OrderBy "receivedDateTime desc"

        if (-not $messages -or $messages.Count -eq 0) {
            Write-Host "No emails found." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        Write-Host "Found $($messages.Count) email(s)" -ForegroundColor $Global:ColorScheme.Success
        Start-Sleep -Seconds 1

        # Prepare messages for display
        $messagesForView = @()
        foreach ($msg in $messages) {
            $messagesForView += [PSCustomObject]@{
                Id                 = $msg.Id
                ReceivedDateTime   = $msg.ReceivedDateTime
                Subject            = $msg.Subject
                SenderName         = if ($msg.From -and $msg.From.EmailAddress) { $msg.From.EmailAddress.Name } else { "N/A" }
                SenderEmailAddress = if ($msg.From -and $msg.From.EmailAddress) { $msg.From.EmailAddress.Address } else { "N/A" }
                Size               = if ($msg.PSObject.Properties['Size']) { $msg.Size } else { 0 }
                HasAttachments     = if ($msg.PSObject.Properties['HasAttachments']) { $msg.HasAttachments } else { $false }
                BodyPreview        = if ($msg.PSObject.Properties['BodyPreview']) { $msg.BodyPreview } else { "" }
            }
        }

        # Display using standardized email list view
        Show-StandardizedEmailListView -UserEmail $UserEmail `
                                       -Messages $messagesForView `
                                       -Title "Recent $Count Emails" `
                                       -AllowActions $true `
                                       -ViewName "RecentEmails"
    }
    catch {
        Write-Error "Error fetching recent emails: $($_.Exception.Message)"
        Write-Host "`nAn error occurred while fetching recent emails." -ForegroundColor $Global:ColorScheme.Error
        Write-Host $_.Exception.Message -ForegroundColor $Global:ColorScheme.Error
        Read-Host "Press Enter to continue"
    }
}

# Function: Show-SenderEmailsMenu
function Show-SenderEmailsMenu {
    <#
    .SYNOPSIS
        Shows menu to select a sender and view their emails
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Clear-Host
        Write-Host "`nManage Emails from Specific Sender" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Get sender cache
        $cache = Get-SenderCache

        if ($null -eq $cache -or $cache.Count -eq 0) {
            Write-Host "No sender cache found. Please rebuild the cache first." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        # Convert cache to array for selection
        $senderList = @()
        foreach ($domainKey in $cache.Keys) {
            $senderData = $cache[$domainKey]
            $senderList += [PSCustomObject]@{
                Domain      = $domainKey
                Count       = $senderData.Count
                DisplayText = "$($senderData.Count.ToString().PadLeft(6)) emails | $domainKey"
            }
        }

        # Sort by count descending
        $senderList = $senderList | Sort-Object Count -Descending

        # Show selectable list
        $selected = Show-SelectableList -Title "Select Sender to Manage" `
                                        -Items $senderList `
                                        -DisplayProperty "DisplayText" `
                                        -PageSize 30

        if ($selected) {
            # Get emails from selected sender
            $senderData = $cache[$selected.Domain]

            if ($senderData -and $senderData.Messages) {
                Write-Host "`nLoading emails from: $($selected.Domain)" -ForegroundColor $Global:ColorScheme.Info
                Start-Sleep -Seconds 1

                # Prepare messages for display
                $messagesForView = @()
                foreach ($msg in $senderData.Messages) {
                    # Cache messages use MessageId property, not Id
                    # Handle both hashtables (from cache) and PSCustomObjects
                    $msgId = $null
                    if ($msg.MessageId) {
                        $msgId = $msg.MessageId
                    } elseif ($msg.Id) {
                        $msgId = $msg.Id
                    }

                    $messagesForView += [PSCustomObject]@{
                        Id                 = $msgId
                        MessageId          = $msgId  # Also add MessageId for compatibility
                        ReceivedDateTime   = $msg.ReceivedDateTime
                        Subject            = $msg.Subject
                        SenderName         = if ($msg.SenderName) { $msg.SenderName } else { "N/A" }
                        SenderEmailAddress = if ($msg.SenderEmailAddress) { $msg.SenderEmailAddress } else { "N/A" }
                        Size               = if ($msg.Size) { $msg.Size } else { 0 }
                        HasAttachments     = if ($msg.HasAttachments) { $msg.HasAttachments } else { $false }
                        BodyPreview        = if ($msg.BodyPreview) { $msg.BodyPreview } else { "" }
                    }
                }

                # Display using standardized email list view
                Show-StandardizedEmailListView -UserEmail $UserEmail `
                                               -Messages $messagesForView `
                                               -Title "Emails from: $($selected.Domain)" `
                                               -AllowActions $true `
                                               -ViewName "SenderEmails_$($selected.Domain)"
            } else {
                Write-Host "No emails found for this sender." -ForegroundColor $Global:ColorScheme.Warning
                Read-Host "Press Enter to continue"
            }
        }
    }
    catch {
        Write-Error "Error showing sender emails: $($_.Exception.Message)"
        Write-Host "`nAn error occurred." -ForegroundColor $Global:ColorScheme.Error
        Write-Host $_.Exception.Message -ForegroundColor $Global:ColorScheme.Error
        Read-Host "Press Enter to continue"
    }
}

# Export functions
Export-ModuleMember -Function Invoke-EmailSearch, Show-RecentEmails, Show-SenderEmailsMenu
