<#
.SYNOPSIS
    Advanced Filter Engine for MailCleanBuddy
.DESCRIPTION
    Provides advanced filtering capabilities with AND/OR logic, saved filters, and filter combinations
#>

# Script-level saved filters storage
$Script:SavedFiltersPath = $null

<#
.SYNOPSIS
    Initializes the filter engine
#>
function Initialize-FilterEngine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
        $homeDir = if ($IsWindows -or $null -eq $IsWindows) { $env:USERPROFILE } else { $env:HOME }
        $filterDir = Join-Path $homeDir ".mailcleanbuddy"

        if (-not (Test-Path $filterDir)) {
            New-Item -Path $filterDir -ItemType Directory -Force | Out-Null
        }

        $Script:SavedFiltersPath = Join-Path $filterDir "saved_filters_$sanitizedEmail.json"

        if (-not (Test-Path $Script:SavedFiltersPath)) {
            $initialData = @{
                Version = "1.0"
                Filters = @()
                LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
            $initialData | ConvertTo-Json -Depth 10 | Set-Content -Path $Script:SavedFiltersPath -Encoding UTF8
        }

        return $true
    } catch {
        Write-Warning "Failed to initialize filter engine: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Creates a new filter object
#>
function New-EmailFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [bool]$HasAttachments = $null,

        [Parameter(Mandatory = $false)]
        [bool]$IsRead = $null,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Low", "Normal", "High", $null)]
        [string]$Importance = $null,

        [Parameter(Mandatory = $false)]
        [int]$MinSize = 0,

        [Parameter(Mandatory = $false)]
        [int]$MaxSize = 0,

        [Parameter(Mandatory = $false)]
        [datetime]$DateFrom = $null,

        [Parameter(Mandatory = $false)]
        [datetime]$DateTo = $null,

        [Parameter(Mandatory = $false)]
        [string]$SenderDomain = $null,

        [Parameter(Mandatory = $false)]
        [string]$SubjectContains = $null,

        [Parameter(Mandatory = $false)]
        [string[]]$Categories = @(),

        [Parameter(Mandatory = $false)]
        [ValidateSet("And", "Or")]
        [string]$LogicOperator = "And"
    )

    return [PSCustomObject]@{
        Name = $Name
        HasAttachments = $HasAttachments
        IsRead = $IsRead
        Importance = $Importance
        MinSize = $MinSize
        MaxSize = $MaxSize
        DateFrom = $DateFrom
        DateTo = $DateTo
        SenderDomain = $SenderDomain
        SubjectContains = $SubjectContains
        Categories = $Categories
        LogicOperator = $LogicOperator
        CreatedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
}

<#
.SYNOPSIS
    Applies a filter to a collection of messages
#>
function Invoke-EmailFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Messages,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Filter
    )

    try {
        $results = @()

        foreach ($message in $Messages) {
            $matches = Test-MessageMatchesFilter -Message $message -Filter $Filter

            if ($matches) {
                $results += $message
            }
        }

        return $results
    } catch {
        Write-Error "Error applying filter: $($_.Exception.Message)"
        return @()
    }
}

<#
.SYNOPSIS
    Tests if a message matches a filter
#>
function Test-MessageMatchesFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Message,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Filter
    )

    $conditions = @()

    # Check HasAttachments
    if ($null -ne $Filter.HasAttachments) {
        $conditions += ($Message.HasAttachments -eq $Filter.HasAttachments)
    }

    # Check IsRead
    if ($null -ne $Filter.IsRead) {
        $isRead = if ($Message.PSObject.Properties['IsRead']) { $Message.IsRead } else { $false }
        $conditions += ($isRead -eq $Filter.IsRead)
    }

    # Check Importance
    if ($Filter.Importance) {
        $importance = if ($Message.PSObject.Properties['Importance']) { $Message.Importance } else { "Normal" }
        $conditions += ($importance -eq $Filter.Importance)
    }

    # Check Size
    if ($Filter.MinSize -gt 0) {
        $size = if ($Message.PSObject.Properties['Size']) { $Message.Size } else { 0 }
        $conditions += ($size -ge $Filter.MinSize)
    }

    if ($Filter.MaxSize -gt 0) {
        $size = if ($Message.PSObject.Properties['Size']) { $Message.Size } else { 0 }
        $conditions += ($size -le $Filter.MaxSize)
    }

    # Check Date Range
    if ($Filter.DateFrom) {
        $receivedDate = ConvertTo-SafeDateTime -DateTimeValue $Message.ReceivedDateTime
        if ($receivedDate) {
            $conditions += ($receivedDate -ge $Filter.DateFrom)
        }
    }

    if ($Filter.DateTo) {
        $receivedDate = ConvertTo-SafeDateTime -DateTimeValue $Message.ReceivedDateTime
        if ($receivedDate) {
            $conditions += ($receivedDate -le $Filter.DateTo)
        }
    }

    # Check Sender Domain
    if ($Filter.SenderDomain) {
        $senderEmail = if ($Message.PSObject.Properties['SenderEmailAddress']) {
            $Message.SenderEmailAddress
        } else {
            ""
        }
        $conditions += ($senderEmail -like "*$($Filter.SenderDomain)*")
    }

    # Check Subject Contains
    if ($Filter.SubjectContains) {
        $subject = if ($Message.PSObject.Properties['Subject']) { $Message.Subject } else { "" }
        $conditions += ($subject -like "*$($Filter.SubjectContains)*")
    }

    # Check Categories
    if ($Filter.Categories -and $Filter.Categories.Count -gt 0) {
        $msgCategories = if ($Message.PSObject.Properties['Categories']) { $Message.Categories } else { @() }
        $categoryMatch = $false
        foreach ($cat in $Filter.Categories) {
            if ($msgCategories -contains $cat) {
                $categoryMatch = $true
                break
            }
        }
        $conditions += $categoryMatch
    }

    # Apply logic operator
    if ($conditions.Count -eq 0) {
        return $true  # No filters applied, match all
    }

    if ($Filter.LogicOperator -eq "Or") {
        return ($conditions -contains $true)
    } else {
        return ($conditions -notcontains $false)
    }
}

<#
.SYNOPSIS
    Combines multiple filters
#>
function Invoke-CombinedFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Messages,

        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$Filters,

        [Parameter(Mandatory = $false)]
        [ValidateSet("And", "Or")]
        [string]$CombineOperator = "And"
    )

    try {
        if ($Filters.Count -eq 0) {
            return $Messages
        }

        if ($CombineOperator -eq "And") {
            # Messages must match ALL filters
            $results = $Messages
            foreach ($filter in $Filters) {
                $results = Invoke-EmailFilter -Messages $results -Filter $filter
            }
            return $results
        } else {
            # Messages must match ANY filter (union)
            $resultSet = @{}
            foreach ($filter in $Filters) {
                $filtered = Invoke-EmailFilter -Messages $Messages -Filter $filter
                foreach ($msg in $filtered) {
                    $msgId = if ($msg.Id) { $msg.Id } elseif ($msg.MessageId) { $msg.MessageId } else { $null }
                    if ($msgId -and -not $resultSet.ContainsKey($msgId)) {
                        $resultSet[$msgId] = $msg
                    }
                }
            }
            return $resultSet.Values
        }
    } catch {
        Write-Error "Error applying combined filter: $($_.Exception.Message)"
        return @()
    }
}

<#
.SYNOPSIS
    Saves a filter for future use
#>
function Save-EmailFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Filter
    )

    try {
        if (-not $Script:SavedFiltersPath) {
            throw "Filter engine not initialized"
        }

        $data = Get-Content -Path $Script:SavedFiltersPath -Raw | ConvertFrom-Json -AsHashtable

        # Check for duplicate names
        $existingIndex = -1
        for ($i = 0; $i -lt $data.Filters.Count; $i++) {
            if ($data.Filters[$i].Name -eq $Filter.Name) {
                $existingIndex = $i
                break
            }
        }

        if ($existingIndex -ge 0) {
            # Update existing filter
            $data.Filters[$existingIndex] = $Filter
            Write-Verbose "Updated existing filter: $($Filter.Name)"
        } else {
            # Add new filter
            $data.Filters += $Filter
            Write-Verbose "Added new filter: $($Filter.Name)"
        }

        $data.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $Script:SavedFiltersPath -Encoding UTF8

        return $true
    } catch {
        Write-Error "Failed to save filter: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Gets all saved filters
#>
function Get-SavedFilters {
    [CmdletBinding()]
    param()

    try {
        if (-not $Script:SavedFiltersPath -or -not (Test-Path $Script:SavedFiltersPath)) {
            return @()
        }

        $data = Get-Content -Path $Script:SavedFiltersPath -Raw | ConvertFrom-Json
        return $data.Filters
    } catch {
        Write-Warning "Failed to load saved filters: $($_.Exception.Message)"
        return @()
    }
}

<#
.SYNOPSIS
    Gets a saved filter by name
#>
function Get-SavedFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $filters = Get-SavedFilters
    return $filters | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
}

<#
.SYNOPSIS
    Removes a saved filter
#>
function Remove-SavedFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        if (-not $Script:SavedFiltersPath) {
            throw "Filter engine not initialized"
        }

        $data = Get-Content -Path $Script:SavedFiltersPath -Raw | ConvertFrom-Json -AsHashtable
        $data.Filters = $data.Filters | Where-Object { $_.Name -ne $Name }
        $data.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $Script:SavedFiltersPath -Encoding UTF8

        return $true
    } catch {
        Write-Error "Failed to remove filter: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Interactive filter builder UI
#>
function Show-FilterBuilder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Clear-Host
        Write-Host "`n🔍 Advanced Filter Builder" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Initialize-FilterEngine -UserEmail $UserEmail | Out-Null

        # Build filter interactively
        Write-Host "Let's build a custom filter. Leave blank to skip a criterion." -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        $filterName = Read-Host "Filter name (required)"
        if ([string]::IsNullOrWhiteSpace($filterName)) {
            Write-Host "Filter name is required." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return $null
        }

        Write-Host ""
        Write-Host "📎 Attachment Filter" -ForegroundColor $Global:ColorScheme.SectionHeader
        $hasAttach = Read-Host "Has attachments? (yes/no/skip)"
        $hasAttachments = $null
        if ($hasAttach -eq "yes" -or $hasAttach -eq "y") { $hasAttachments = $true }
        elseif ($hasAttach -eq "no" -or $hasAttach -eq "n") { $hasAttachments = $false }

        Write-Host ""
        Write-Host "👁️ Read Status Filter" -ForegroundColor $Global:ColorScheme.SectionHeader
        $readStatus = Read-Host "Is read? (yes/no/skip)"
        $isRead = $null
        if ($readStatus -eq "yes" -or $readStatus -eq "y") { $isRead = $true }
        elseif ($readStatus -eq "no" -or $readStatus -eq "n") { $isRead = $false }

        Write-Host ""
        Write-Host "⚠️ Importance Filter" -ForegroundColor $Global:ColorScheme.SectionHeader
        $importance = Read-Host "Importance (Low/Normal/High/skip)"
        if ($importance -notin @("Low", "Normal", "High")) { $importance = $null }

        Write-Host ""
        Write-Host "📏 Size Filter" -ForegroundColor $Global:ColorScheme.SectionHeader
        $minSizeInput = Read-Host "Minimum size in KB (or skip)"
        $minSize = 0
        if ($minSizeInput -match '^\d+$') { $minSize = [int]$minSizeInput * 1KB }

        $maxSizeInput = Read-Host "Maximum size in KB (or skip)"
        $maxSize = 0
        if ($maxSizeInput -match '^\d+$') { $maxSize = [int]$maxSizeInput * 1KB }

        Write-Host ""
        Write-Host "📅 Date Range Filter" -ForegroundColor $Global:ColorScheme.SectionHeader
        $dateFromInput = Read-Host "From date (yyyy-MM-dd or skip)"
        $dateFrom = $null
        if (-not [string]::IsNullOrWhiteSpace($dateFromInput)) {
            try { $dateFrom = [datetime]::Parse($dateFromInput) }
            catch { Write-Host "Invalid date format, skipping" -ForegroundColor $Global:ColorScheme.Warning }
        }

        $dateToInput = Read-Host "To date (yyyy-MM-dd or skip)"
        $dateTo = $null
        if (-not [string]::IsNullOrWhiteSpace($dateToInput)) {
            try { $dateTo = [datetime]::Parse($dateToInput) }
            catch { Write-Host "Invalid date format, skipping" -ForegroundColor $Global:ColorScheme.Warning }
        }

        Write-Host ""
        Write-Host "📧 Content Filters" -ForegroundColor $Global:ColorScheme.SectionHeader
        $senderDomain = Read-Host "Sender domain (e.g., gmail.com or skip)"
        if ([string]::IsNullOrWhiteSpace($senderDomain)) { $senderDomain = $null }

        $subjectContains = Read-Host "Subject contains (or skip)"
        if ([string]::IsNullOrWhiteSpace($subjectContains)) { $subjectContains = $null }

        # Create filter
        $filter = New-EmailFilter -Name $filterName `
                                  -HasAttachments $hasAttachments `
                                  -IsRead $isRead `
                                  -Importance $importance `
                                  -MinSize $minSize `
                                  -MaxSize $maxSize `
                                  -DateFrom $dateFrom `
                                  -DateTo $dateTo `
                                  -SenderDomain $senderDomain `
                                  -SubjectContains $subjectContains

        Write-Host ""
        Write-Host "✅ Filter created successfully!" -ForegroundColor $Global:ColorScheme.Success

        $save = Read-Host "Save this filter? (yes/no)"
        if ($save -eq "yes" -or $save -eq "y") {
            if (Save-EmailFilter -Filter $filter) {
                Write-Host "Filter saved!" -ForegroundColor $Global:ColorScheme.Success
            }
        }

        return $filter
    } catch {
        Write-Error "Error in filter builder: $($_.Exception.Message)"
        Read-Host "Press Enter to continue"
        return $null
    }
}

<#
.SYNOPSIS
    Shows the filter management menu
#>
function Show-FilterManagement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Initialize-FilterEngine -UserEmail $UserEmail | Out-Null

        while ($true) {
            Clear-Host
            Write-Host "`n🔍 Filter Management" -ForegroundColor $Global:ColorScheme.Highlight
            Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""

            $savedFilters = Get-SavedFilters

            Write-Host "📋 Saved Filters ($($savedFilters.Count))" -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ""

            if ($savedFilters.Count -eq 0) {
                Write-Host "  No saved filters yet." -ForegroundColor $Global:ColorScheme.Muted
            } else {
                $index = 1
                foreach ($filter in $savedFilters) {
                    Write-Host "  $index. " -NoNewline -ForegroundColor $Global:ColorScheme.Muted
                    Write-Host "$($filter.Name)" -ForegroundColor $Global:ColorScheme.Value

                    # Show summary
                    $summary = @()
                    if ($null -ne $filter.HasAttachments) { $summary += "Attachments: $($filter.HasAttachments)" }
                    if ($null -ne $filter.IsRead) { $summary += "Read: $($filter.IsRead)" }
                    if ($filter.Importance) { $summary += "Importance: $($filter.Importance)" }
                    if ($filter.MinSize -gt 0) { $summary += "Min: $([Math]::Round($filter.MinSize/1KB))KB" }
                    if ($filter.SenderDomain) { $summary += "From: $($filter.SenderDomain)" }

                    if ($summary.Count -gt 0) {
                        Write-Host "     " -NoNewline
                        Write-Host ($summary -join ", ") -ForegroundColor $Global:ColorScheme.Muted
                    }

                    $index++
                }
            }

            Write-Host ""
            Write-Host "Actions:" -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  [N] Create new filter" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [D] Delete a filter" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [Q] Back" -ForegroundColor $Global:ColorScheme.Info
            Write-Host ""

            $choice = Read-Host "Select action"

            switch ($choice.ToUpper()) {
                "N" {
                    Show-FilterBuilder -UserEmail $UserEmail
                    Read-Host "Press Enter to continue"
                }
                "D" {
                    if ($savedFilters.Count -gt 0) {
                        $deleteNum = Read-Host "Enter filter number to delete"
                        if ($deleteNum -match '^\d+$' -and [int]$deleteNum -ge 1 -and [int]$deleteNum -le $savedFilters.Count) {
                            $filterToDelete = $savedFilters[[int]$deleteNum - 1]
                            $confirm = Read-Host "Delete filter '$($filterToDelete.Name)'? (yes/no)"
                            if ($confirm -eq "yes" -or $confirm -eq "y") {
                                Remove-SavedFilter -Name $filterToDelete.Name
                                Write-Host "Filter deleted." -ForegroundColor $Global:ColorScheme.Success
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                }
                "Q" {
                    return
                }
            }
        }
    } catch {
        Write-Error "Error in filter management: $($_.Exception.Message)"
        Read-Host "Press Enter to continue"
    }
}

# Export functions
Export-ModuleMember -Function Initialize-FilterEngine, New-EmailFilter, Invoke-EmailFilter, `
                              Test-MessageMatchesFilter, Invoke-CombinedFilter, Save-EmailFilter, `
                              Get-SavedFilters, Get-SavedFilter, Remove-SavedFilter, `
                              Show-FilterBuilder, Show-FilterManagement
