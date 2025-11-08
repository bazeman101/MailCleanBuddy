<#
.SYNOPSIS
    Advanced Email Search module for MailCleanBuddy
.DESCRIPTION
    Provides advanced search capabilities with regex, filters, saved queries, and search analytics.
#>

# Import dependencies

# Saved queries database path
$script:SavedQueriesPath = $null

# Function: Initialize-AdvancedSearch
function Initialize-AdvancedSearch {
    <#
    .SYNOPSIS
        Initializes advanced search database
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
    $script:SavedQueriesPath = Join-Path $PSScriptRoot "..\..\saved_queries_$sanitizedEmail.json"

    if (-not (Test-Path $script:SavedQueriesPath)) {
        $initialData = @{
            Queries = @()
            SearchHistory = @()
        }
        $initialData | ConvertTo-Json -Depth 10 | Set-Content -Path $script:SavedQueriesPath -Encoding UTF8
    }
}

# Function: Invoke-AdvancedSearch
function Invoke-AdvancedSearch {
    <#
    .SYNOPSIS
        Performs advanced email search
    .PARAMETER UserEmail
        User email address
    .PARAMETER SearchCriteria
        Search criteria object
    .OUTPUTS
        Array of matching emails
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$SearchCriteria
    )

    try {
        $results = @()
        $cache = Get-SenderCache

        Write-Host ""
        Write-Host (Get-LocalizedString "advSearch_searching") -ForegroundColor $Global:ColorScheme.Info

        $totalEmails = 0
        foreach ($domain in $cache.Keys) {
            $totalEmails += $cache[$domain].Messages.Count
        }

        $processed = 0
        $progressId = 1

        foreach ($domain in $cache.Keys) {
            foreach ($message in $cache[$domain].Messages) {
                $processed++

                if ($processed % 50 -eq 0) {
                    Write-Progress -Id $progressId -Activity (Get-LocalizedString "advSearch_progressActivity") `
                        -Status (Get-LocalizedString "advSearch_progressStatus" -FormatArgs @($processed, $totalEmails)) `
                        -PercentComplete (($processed / $totalEmails) * 100)
                }

                $matches = $true

                # Text search (subject, body preview, sender)
                if ($SearchCriteria.SearchText) {
                    $searchText = $SearchCriteria.SearchText
                    $isRegex = $SearchCriteria.UseRegex

                    $textFields = @(
                        $message.Subject,
                        $message.BodyPreview,
                        $message.SenderEmailAddress,
                        $message.SenderName
                    )

                    $textMatch = $false
                    foreach ($field in $textFields) {
                        if ($field) {
                            if ($isRegex) {
                                if ($field -match $searchText) {
                                    $textMatch = $true
                                    break
                                }
                            } else {
                                if ($field -like "*$searchText*") {
                                    $textMatch = $true
                                    break
                                }
                            }
                        }
                    }

                    if (-not $textMatch) {
                        $matches = $false
                    }
                }

                # Date range filter
                if ($matches -and $SearchCriteria.DateFrom) {
                    $receivedDate = ConvertTo-SafeDateTime -DateTimeValue $message.ReceivedDateTime
                    if ($receivedDate -lt $SearchCriteria.DateFrom) {
                        $matches = $false
                    }
                }

                if ($matches -and $SearchCriteria.DateTo) {
                    $receivedDate = ConvertTo-SafeDateTime -DateTimeValue $message.ReceivedDateTime
                    if ($receivedDate -gt $SearchCriteria.DateTo) {
                        $matches = $false
                    }
                }

                # Size filter
                if ($matches -and $SearchCriteria.MinSize) {
                    if ($message.Size -lt $SearchCriteria.MinSize) {
                        $matches = $false
                    }
                }

                if ($matches -and $SearchCriteria.MaxSize) {
                    if ($message.Size -gt $SearchCriteria.MaxSize) {
                        $matches = $false
                    }
                }

                # Has attachments filter
                if ($matches -and $SearchCriteria.HasAttachments -ne $null) {
                    if ($SearchCriteria.HasAttachments -and -not $message.HasAttachments) {
                        $matches = $false
                    }
                    if (-not $SearchCriteria.HasAttachments -and $message.HasAttachments) {
                        $matches = $false
                    }
                }

                # Is read filter
                if ($matches -and $SearchCriteria.IsRead -ne $null) {
                    if ($SearchCriteria.IsRead -and -not $message.IsRead) {
                        $matches = $false
                    }
                    if (-not $SearchCriteria.IsRead -and $message.IsRead) {
                        $matches = $false
                    }
                }

                # Sender filter
                if ($matches -and $SearchCriteria.SenderDomain) {
                    if ($message.SenderEmailAddress -notlike "*$($SearchCriteria.SenderDomain)*") {
                        $matches = $false
                    }
                }

                if ($matches) {
                    $results += $message
                }
            }
        }

        Write-Progress -Id $progressId -Activity (Get-LocalizedString "advSearch_progressActivity") -Completed

        # Save to search history
        Save-SearchToHistory -SearchCriteria $SearchCriteria -ResultCount $results.Count

        return $results
    }
    catch {
        Write-Error "Error performing advanced search: $($_.Exception.Message)"
        return @()
    }
}

# Function: Show-AdvancedSearch
function Show-AdvancedSearch {
    <#
    .SYNOPSIS
        Interactive advanced search interface
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Initialize-AdvancedSearch -UserEmail $UserEmail

        Clear-Host

        $title = Get-LocalizedString "advSearch_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host (Get-LocalizedString "advSearch_description") -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        while ($true) {
            Write-Host (Get-LocalizedString "advSearch_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  1. $(Get-LocalizedString 'advSearch_newSearch')" -ForegroundColor Green
            Write-Host "  2. $(Get-LocalizedString 'advSearch_loadSavedQuery')" -ForegroundColor Cyan
            Write-Host "  3. $(Get-LocalizedString 'advSearch_viewHistory')" -ForegroundColor Yellow
            Write-Host "  4. $(Get-LocalizedString 'advSearch_manageSavedQueries')" -ForegroundColor Magenta
            Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
            Write-Host ""

            $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

            switch ($choice.ToUpper()) {
                "1" {
                    # New search
                    $criteria = Build-SearchCriteria
                    if ($criteria) {
                        $results = Invoke-AdvancedSearch -UserEmail $UserEmail -SearchCriteria $criteria
                        Show-SearchResults -UserEmail $UserEmail -Results $results -Criteria $criteria
                    }
                }
                "2" {
                    # Load saved query
                    $query = Select-SavedQuery
                    if ($query) {
                        $results = Invoke-AdvancedSearch -UserEmail $UserEmail -SearchCriteria $query.Criteria
                        Show-SearchResults -UserEmail $UserEmail -Results $results -Criteria $query.Criteria
                    }
                }
                "3" {
                    # View history
                    Show-SearchHistory
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "4" {
                    # Manage saved queries
                    Manage-SavedQueries
                }
                "Q" {
                    return
                }
                default {
                    Write-Host (Get-LocalizedString "unsubscribe_invalidChoice") -ForegroundColor $Global:ColorScheme.Warning
                }
            }

            Write-Host ""
        }
    }
    catch {
        Write-Error "Error in advanced search: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Build-SearchCriteria
function Build-SearchCriteria {
    <#
    .SYNOPSIS
        Builds search criteria from user input
    .OUTPUTS
        Search criteria object
    #>
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host (Get-LocalizedString "advSearch_buildCriteria") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 80) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    $criteria = [PSCustomObject]@{
        SearchText = $null
        UseRegex = $false
        DateFrom = $null
        DateTo = $null
        MinSize = $null
        MaxSize = $null
        HasAttachments = $null
        IsRead = $null
        SenderDomain = $null
    }

    # Text search
    $searchText = Read-Host (Get-LocalizedString "advSearch_enterSearchText")
    if (-not [string]::IsNullOrWhiteSpace($searchText)) {
        $criteria.SearchText = $searchText

        $useRegex = Read-Host (Get-LocalizedString "advSearch_useRegex")
        $criteria.UseRegex = ($useRegex -eq "yes" -or $useRegex -eq "ja" -or $useRegex -eq "oui" -or $useRegex -eq "y")
    }

    # Date range
    $dateFrom = Read-Host (Get-LocalizedString "advSearch_enterDateFrom")
    if (-not [string]::IsNullOrWhiteSpace($dateFrom)) {
        try {
            $criteria.DateFrom = [DateTime]::Parse($dateFrom)
        } catch {
            Write-Host (Get-LocalizedString "advSearch_invalidDate") -ForegroundColor $Global:ColorScheme.Warning
        }
    }

    $dateTo = Read-Host (Get-LocalizedString "advSearch_enterDateTo")
    if (-not [string]::IsNullOrWhiteSpace($dateTo)) {
        try {
            $criteria.DateTo = [DateTime]::Parse($dateTo)
        } catch {
            Write-Host (Get-LocalizedString "advSearch_invalidDate") -ForegroundColor $Global:ColorScheme.Warning
        }
    }

    # Size filters
    $minSize = Read-Host (Get-LocalizedString "advSearch_enterMinSize")
    if (-not [string]::IsNullOrWhiteSpace($minSize)) {
        $criteria.MinSize = [int]$minSize * 1KB
    }

    $maxSize = Read-Host (Get-LocalizedString "advSearch_enterMaxSize")
    if (-not [string]::IsNullOrWhiteSpace($maxSize)) {
        $criteria.MaxSize = [int]$maxSize * 1KB
    }

    # Has attachments
    $hasAttach = Read-Host (Get-LocalizedString "advSearch_hasAttachments")
    if (-not [string]::IsNullOrWhiteSpace($hasAttach)) {
        $criteria.HasAttachments = ($hasAttach -eq "yes" -or $hasAttach -eq "ja" -or $hasAttach -eq "oui" -or $hasAttach -eq "y")
    }

    # Is read
    $isRead = Read-Host (Get-LocalizedString "advSearch_isRead")
    if (-not [string]::IsNullOrWhiteSpace($isRead)) {
        $criteria.IsRead = ($isRead -eq "yes" -or $isRead -eq "ja" -or $isRead -eq "oui" -or $isRead -eq "y")
    }

    # Sender domain
    $senderDomain = Read-Host (Get-LocalizedString "advSearch_enterSenderDomain")
    if (-not [string]::IsNullOrWhiteSpace($senderDomain)) {
        $criteria.SenderDomain = $senderDomain
    }

    # Ask if user wants to save this query
    $saveName = Read-Host (Get-LocalizedString "advSearch_saveQueryPrompt")
    if (-not [string]::IsNullOrWhiteSpace($saveName)) {
        Save-SearchQuery -Name $saveName -Criteria $criteria
        Write-Host (Get-LocalizedString "advSearch_querySaved") -ForegroundColor $Global:ColorScheme.Success
    }

    return $criteria
}

# Function: Show-SearchResults
function Show-SearchResults {
    <#
    .SYNOPSIS
        Displays search results
    .PARAMETER UserEmail
        User email address
    .PARAMETER Results
        Search results
    .PARAMETER Criteria
        Search criteria used
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Results,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Criteria
    )

    Write-Host ""
    Write-Host (Get-LocalizedString "advSearch_resultsFound" -FormatArgs @($Results.Count)) -ForegroundColor $Global:ColorScheme.Success
    Write-Host ""

    if ($Results.Count -eq 0) {
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    # Show results using standardized list view
    $listParams = @{
        UserEmail = $UserEmail
        Messages = $Results
        Title = "Advanced Search Results"
        AllowActions = $true
        ViewName = "AdvancedSearchResults"
    }

    Show-StandardizedEmailListView @listParams
}

# Function: Save-SearchQuery
function Save-SearchQuery {
    <#
    .SYNOPSIS
        Saves a search query
    .PARAMETER Name
        Query name
    .PARAMETER Criteria
        Search criteria
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Criteria
    )

    $data = Get-Content -Path $script:SavedQueriesPath -Raw | ConvertFrom-Json

    $query = [PSCustomObject]@{
        Name = $Name
        Criteria = $Criteria
        CreatedDate = (Get-Date).ToString("o")
        LastUsed = (Get-Date).ToString("o")
        UseCount = 0
    }

    $data.Queries += $query
    $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:SavedQueriesPath -Encoding UTF8
}

# Function: Save-SearchToHistory
function Save-SearchToHistory {
    <#
    .SYNOPSIS
        Saves search to history
    .PARAMETER SearchCriteria
        Search criteria
    .PARAMETER ResultCount
        Number of results
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$SearchCriteria,

        [Parameter(Mandatory = $true)]
        [int]$ResultCount
    )

    try {
        $data = Get-Content -Path $script:SavedQueriesPath -Raw | ConvertFrom-Json

        $historyEntry = [PSCustomObject]@{
            Criteria = $SearchCriteria
            ResultCount = $ResultCount
            Timestamp = (Get-Date).ToString("o")
        }

        $data.SearchHistory += $historyEntry

        # Keep only last 50 searches
        if ($data.SearchHistory.Count -gt 50) {
            $data.SearchHistory = $data.SearchHistory | Select-Object -Last 50
        }

        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:SavedQueriesPath -Encoding UTF8
    }
    catch {
        Write-Warning "Could not save search to history: $($_.Exception.Message)"
    }
}

# Function: Select-SavedQuery
function Select-SavedQuery {
    <#
    .SYNOPSIS
        Selects a saved query
    .OUTPUTS
        Selected query object
    #>
    [CmdletBinding()]
    param()

    $data = Get-Content -Path $script:SavedQueriesPath -Raw | ConvertFrom-Json

    if ($data.Queries.Count -eq 0) {
        Write-Host "`n$(Get-LocalizedString 'advSearch_noSavedQueries')" -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return $null
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "advSearch_savedQueriesList") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 80) -ForegroundColor $Global:ColorScheme.Border

    $index = 1
    foreach ($query in $data.Queries) {
        Write-Host "  [$index] " -NoNewline -ForegroundColor $Global:ColorScheme.Muted
        Write-Host "$($query.Name)" -NoNewline -ForegroundColor $Global:ColorScheme.Value
        Write-Host " (used $($query.UseCount) times)" -ForegroundColor $Global:ColorScheme.Muted
        $index++
    }

    Write-Host ""
    $selection = Read-Host (Get-LocalizedString "advSearch_selectQuery")

    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $data.Queries.Count) {
        $selectedQuery = $data.Queries[[int]$selection - 1]
        $selectedQuery.LastUsed = (Get-Date).ToString("o")
        $selectedQuery.UseCount++

        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:SavedQueriesPath -Encoding UTF8

        return $selectedQuery
    }

    return $null
}

# Function: Show-SearchHistory
function Show-SearchHistory {
    <#
    .SYNOPSIS
        Shows search history
    #>
    [CmdletBinding()]
    param()

    $data = Get-Content -Path $script:SavedQueriesPath -Raw | ConvertFrom-Json

    if ($data.SearchHistory.Count -eq 0) {
        Write-Host "`n$(Get-LocalizedString 'advSearch_noHistory')" -ForegroundColor $Global:ColorScheme.Warning
        return
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "advSearch_historyTitle" -FormatArgs @($data.SearchHistory.Count)) -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 80) -ForegroundColor $Global:ColorScheme.Border

    $recent = $data.SearchHistory | Select-Object -Last 20 | Sort-Object { ConvertTo-SafeDateTime -DateTimeValue $_.Timestamp } -Descending

    foreach ($entry in $recent) {
        $timestamp = ConvertTo-SafeDateTime -DateTimeValue $entry.Timestamp.ToString('yyyy-MM-dd HH:mm')
        Write-Host "  [$timestamp] " -NoNewline -ForegroundColor $Global:ColorScheme.Muted

        if ($entry.Criteria.SearchText) {
            Write-Host "$($entry.Criteria.SearchText)" -NoNewline -ForegroundColor $Global:ColorScheme.Value
        } else {
            Write-Host "(filtered search)" -NoNewline -ForegroundColor $Global:ColorScheme.Value
        }

        Write-Host " → $($entry.ResultCount) results" -ForegroundColor $Global:ColorScheme.Info
    }
}

# Function: Manage-SavedQueries
function Manage-SavedQueries {
    <#
    .SYNOPSIS
        Manages saved queries
    #>
    [CmdletBinding()]
    param()

    $data = Get-Content -Path $script:SavedQueriesPath -Raw | ConvertFrom-Json

    if ($data.Queries.Count -eq 0) {
        Write-Host "`n$(Get-LocalizedString 'advSearch_noSavedQueries')" -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "advSearch_manageQueries") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 80) -ForegroundColor $Global:ColorScheme.Border

    $index = 1
    foreach ($query in $data.Queries) {
        Write-Host "  [$index] $($query.Name)" -ForegroundColor $Global:ColorScheme.Value
        $index++
    }

    Write-Host ""
    $deleteNum = Read-Host (Get-LocalizedString "advSearch_deleteQueryPrompt")

    if ($deleteNum -match '^\d+$' -and [int]$deleteNum -ge 1 -and [int]$deleteNum -le $data.Queries.Count) {
        $queryToDelete = $data.Queries[[int]$deleteNum - 1]

        $confirm = Show-Confirmation -Message (Get-LocalizedString "advSearch_confirmDelete" -FormatArgs @($queryToDelete.Name))

        if ($confirm) {
            $data.Queries = $data.Queries | Where-Object { $_.Name -ne $queryToDelete.Name }
            $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:SavedQueriesPath -Encoding UTF8
            Write-Host (Get-LocalizedString "advSearch_queryDeleted") -ForegroundColor $Global:ColorScheme.Success
        }
    }

    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Export functions
Export-ModuleMember -Function Show-AdvancedSearch, Invoke-AdvancedSearch, Initialize-AdvancedSearch
