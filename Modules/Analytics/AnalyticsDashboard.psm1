<#
.SYNOPSIS
    Analytics Dashboard module for MailCleanBuddy
.DESCRIPTION
    Provides mailbox analytics and insights including statistics, trends,
    top senders, storage analysis, and unsubscribe opportunities.
#>

# Import dependencies

# Function: Get-MailboxStatistics
function Get-MailboxStatistics {
    <#
    .SYNOPSIS
        Calculates comprehensive mailbox statistics from cache
    .DESCRIPTION
        Analyzes cached mailbox data to provide total email count, unique senders,
        total size, average email size, and attachment statistics
    .PARAMETER UserEmail
        Email address of the mailbox being analyzed
    .OUTPUTS
        PSCustomObject with mailbox statistics
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            Write-Warning (Get-LocalizedString "analytics_noCacheData")
            return $null
        }

        $totalEmails = 0
        $totalSize = 0
        $emailsWithAttachments = 0
        $uniqueSenders = $cache.Count
        $allMessages = @()

        # Aggregate data from all domains
        foreach ($domain in $cache.Keys) {
            $senderData = $cache[$domain]
            $totalEmails += $senderData.Count

            foreach ($msg in $senderData.Messages) {
                $allMessages += $msg

                if ($msg.Size -and $msg.Size -gt 0) {
                    $totalSize += $msg.Size
                }

                if ($msg.HasAttachments -eq $true) {
                    $emailsWithAttachments++
                }
            }
        }

        $avgEmailSize = if ($totalEmails -gt 0) { [math]::Round($totalSize / $totalEmails, 2) } else { 0 }
        $totalSizeMB = [math]::Round($totalSize / 1MB, 2)

        # Calculate date range
        $dates = $allMessages | Where-Object { $_.ReceivedDateTime } | ForEach-Object {
            ConvertTo-SafeDateTime -DateTimeValue $_.ReceivedDateTime
        } | Where-Object { $_ -ne [DateTime]::MinValue } | Sort-Object

        $oldestEmail = if ($dates.Count -gt 0) { $dates[0] } else { $null }
        $newestEmail = if ($dates.Count -gt 0) { $dates[-1] } else { $null }

        return [PSCustomObject]@{
            TotalEmails = $totalEmails
            UniqueSenderDomains = $uniqueSenders
            TotalSizeBytes = $totalSize
            TotalSizeMB = $totalSizeMB
            AverageEmailSizeBytes = $avgEmailSize
            EmailsWithAttachments = $emailsWithAttachments
            AttachmentPercentage = if ($totalEmails -gt 0) { [math]::Round(($emailsWithAttachments / $totalEmails) * 100, 1) } else { 0 }
            OldestEmailDate = $oldestEmail
            NewestEmailDate = $newestEmail
            UserEmail = $UserEmail
        }
    }
    catch {
        Write-Error "Error calculating mailbox statistics: $($_.Exception.Message)"
        return $null
    }
}

# Function: Get-TopSenders
function Get-TopSenders {
    <#
    .SYNOPSIS
        Gets top email senders by count
    .DESCRIPTION
        Returns the top N sender domains sorted by email count
    .PARAMETER TopCount
        Number of top senders to return (default: 10)
    .OUTPUTS
        Array of PSCustomObject with sender statistics
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$TopCount = 10
    )

    try {
        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            return @()
        }

        $topSenders = @()

        foreach ($domain in $cache.Keys) {
            $senderData = $cache[$domain]
            $totalSize = 0
            $hasAttachments = 0

            foreach ($msg in $senderData.Messages) {
                if ($msg.Size) { $totalSize += $msg.Size }
                if ($msg.HasAttachments) { $hasAttachments++ }
            }

            $topSenders += [PSCustomObject]@{
                Domain = $domain
                Count = $senderData.Count
                TotalSizeMB = [math]::Round($totalSize / 1MB, 2)
                EmailsWithAttachments = $hasAttachments
            }
        }

        return $topSenders | Sort-Object -Property Count -Descending | Select-Object -First $TopCount
    }
    catch {
        Write-Error "Error getting top senders: $($_.Exception.Message)"
        return @()
    }
}

# Function: Get-MailboxGrowthTrend
function Get-MailboxGrowthTrend {
    <#
    .SYNOPSIS
        Analyzes email volume over time
    .DESCRIPTION
        Groups emails by time period (month) to show mailbox growth trend
    .OUTPUTS
        Array of PSCustomObject with period and email count
    #>
    [CmdletBinding()]
    param()

    try {
        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            return @()
        }

        $allMessages = @()

        foreach ($domain in $cache.Keys) {
            $allMessages += $cache[$domain].Messages
        }

        # Group by month
        $groupedByMonth = $allMessages | Where-Object { $_.ReceivedDateTime } |
            Group-Object -Property {
                $date = ConvertTo-SafeDateTime -DateTimeValue $_.ReceivedDateTime
                if ($date -ne [DateTime]::MinValue) {
                    $date.ToString("yyyy-MM")
                } else {
                    "Unknown"
                }
            } |
            Sort-Object Name

        $trend = @()
        foreach ($group in $groupedByMonth) {
            $trend += [PSCustomObject]@{
                Period = $group.Name
                EmailCount = $group.Count
            }
        }

        return $trend
    }
    catch {
        Write-Error "Error calculating mailbox growth trend: $($_.Exception.Message)"
        return @()
    }
}

# Function: Get-StorageByDomain
function Get-StorageByDomain {
    <#
    .SYNOPSIS
        Calculates storage usage per sender domain
    .DESCRIPTION
        Returns storage breakdown showing which domains consume the most space
    .PARAMETER TopCount
        Number of top storage-consuming domains to return (default: 10)
    .OUTPUTS
        Array of PSCustomObject with domain storage statistics
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$TopCount = 10
    )

    try {
        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            return @()
        }

        $storageByDomain = @()

        foreach ($domain in $cache.Keys) {
            $senderData = $cache[$domain]
            $totalSize = 0

            foreach ($msg in $senderData.Messages) {
                if ($msg.Size) { $totalSize += $msg.Size }
            }

            $storageByDomain += [PSCustomObject]@{
                Domain = $domain
                EmailCount = $senderData.Count
                TotalSizeBytes = $totalSize
                TotalSizeMB = [math]::Round($totalSize / 1MB, 2)
                AvgSizeKB = if ($senderData.Count -gt 0) {
                    [math]::Round($totalSize / 1KB / $senderData.Count, 2)
                } else { 0 }
            }
        }

        return $storageByDomain | Sort-Object -Property TotalSizeBytes -Descending | Select-Object -First $TopCount
    }
    catch {
        Write-Error "Error calculating storage by domain: $($_.Exception.Message)"
        return @()
    }
}

# Function: Get-UnsubscribeOpportunities
function Get-UnsubscribeOpportunities {
    <#
    .SYNOPSIS
        Identifies potential newsletter/marketing senders
    .DESCRIPTION
        Analyzes sender patterns to identify newsletters and marketing emails
        that might be candidates for unsubscribing
    .PARAMETER MinEmailCount
        Minimum number of emails to consider a sender as newsletter (default: 5)
    .OUTPUTS
        Array of PSCustomObject with potential unsubscribe candidates
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$MinEmailCount = 5
    )

    try {
        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            return @()
        }

        # Newsletter/marketing keywords (case-insensitive)
        $newsletterKeywords = @(
            'newsletter', 'nieuwsbrief', 'update', 'digest',
            'notification', 'melding', 'promo', 'aanbieding',
            'deal', 'sale', 'marketing', 'campaign', 'no-reply',
            'noreply', 'donotreply', 'automated', 'auto-'
        )

        $opportunities = @()

        foreach ($domain in $cache.Keys) {
            $senderData = $cache[$domain]

            # Skip if below minimum threshold
            if ($senderData.Count -lt $MinEmailCount) {
                continue
            }

            $score = 0
            $reasons = @()

            # Check domain for newsletter patterns
            foreach ($keyword in $newsletterKeywords) {
                if ($domain -like "*$keyword*") {
                    $score += 2
                    $reasons += "Domain contains '$keyword'"
                    break
                }
            }

            # Check subjects for newsletter patterns
            $newsletterSubjects = 0
            foreach ($msg in $senderData.Messages) {
                if ($msg.Subject) {
                    foreach ($keyword in $newsletterKeywords) {
                        if ($msg.Subject -like "*$keyword*") {
                            $newsletterSubjects++
                            break
                        }
                    }
                }
            }

            if ($newsletterSubjects -gt ($senderData.Count * 0.3)) {
                $score += 3
                $reasons += "Many subjects contain newsletter keywords"
            }

            # High volume sender
            if ($senderData.Count -gt 20) {
                $score += 2
                $reasons += "High volume sender (${0} emails)" -f $senderData.Count
            }

            # Check sender email for patterns
            $sampleMessage = $senderData.Messages | Select-Object -First 1
            if ($sampleMessage.SenderEmailAddress) {
                $senderEmail = $sampleMessage.SenderEmailAddress.ToLower()
                if ($senderEmail -like "*no-reply*" -or $senderEmail -like "*noreply*" -or
                    $senderEmail -like "*donotreply*" -or $senderEmail -like "*newsletter*") {
                    $score += 3
                    $reasons += "Sender email suggests automated/newsletter"
                }
            }

            # If score suggests newsletter/marketing
            if ($score -ge 3) {
                $totalSize = ($senderData.Messages | Measure-Object -Property Size -Sum).Sum

                $opportunities += [PSCustomObject]@{
                    Domain = $domain
                    EmailCount = $senderData.Count
                    Score = $score
                    Reasons = $reasons -join "; "
                    TotalSizeMB = [math]::Round($totalSize / 1MB, 2)
                    SampleSender = $sampleMessage.SenderEmailAddress
                    SampleSubject = $sampleMessage.Subject
                }
            }
        }

        return $opportunities | Sort-Object -Property Score, EmailCount -Descending
    }
    catch {
        Write-Error "Error identifying unsubscribe opportunities: $($_.Exception.Message)"
        return @()
    }
}

# Function: Show-AnalyticsDashboard
function Show-AnalyticsDashboard {
    <#
    .SYNOPSIS
        Displays interactive analytics dashboard
    .DESCRIPTION
        Shows comprehensive mailbox insights including statistics, trends, and recommendations
    .PARAMETER UserEmail
        Email address of the mailbox being analyzed
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Clear-Host

        # Display header
        $title = Get-LocalizedString "analytics_dashboardTitle" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 80) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Get mailbox statistics
        Write-Host (Get-LocalizedString "analytics_calculatingStats") -ForegroundColor $Global:ColorScheme.Info
        $stats = Get-MailboxStatistics -UserEmail $UserEmail

        if (-not $stats) {
            Write-Host (Get-LocalizedString "analytics_noCacheData") -ForegroundColor $Global:ColorScheme.Warning
            Write-Host (Get-LocalizedString "analytics_rebuildCachePrompt") -ForegroundColor $Global:ColorScheme.Info
            Write-Host ""
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Display Overview Section
        Write-Host "`n$(Get-LocalizedString 'analytics_sectionOverview')" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

        Write-Host "  $(Get-LocalizedString 'analytics_totalEmails'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host $stats.TotalEmails -ForegroundColor $Global:ColorScheme.Value

        Write-Host "  $(Get-LocalizedString 'analytics_uniqueSenders'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host $stats.UniqueSenderDomains -ForegroundColor $Global:ColorScheme.Value

        Write-Host "  $(Get-LocalizedString 'analytics_totalStorage'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($stats.TotalSizeMB) MB" -ForegroundColor $Global:ColorScheme.Value

        Write-Host "  $(Get-LocalizedString 'analytics_avgEmailSize'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$([math]::Round($stats.AverageEmailSizeBytes / 1KB, 2)) KB" -ForegroundColor $Global:ColorScheme.Value

        Write-Host "  $(Get-LocalizedString 'analytics_emailsWithAttachments'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host "$($stats.EmailsWithAttachments) ($($stats.AttachmentPercentage)%)" -ForegroundColor $Global:ColorScheme.Value

        if ($stats.OldestEmailDate -and $stats.NewestEmailDate) {
            Write-Host "  $(Get-LocalizedString 'analytics_dateRange'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
            Write-Host "$($stats.OldestEmailDate.ToString('yyyy-MM-dd')) $(Get-LocalizedString 'analytics_to') $($stats.NewestEmailDate.ToString('yyyy-MM-dd'))" -ForegroundColor $Global:ColorScheme.Value
        }

        # Display Top Senders Section
        Write-Host "`n$(Get-LocalizedString 'analytics_sectionTopSenders')" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

        $topSenders = Get-TopSenders -TopCount 10

        if ($topSenders.Count -gt 0) {
            $format = "{0,-4} {1,-45} {2,10} {3,12}"
            Write-Host ($format -f "#", (Get-LocalizedString 'senderOverview_headerDomain'),
                (Get-LocalizedString 'senderOverview_headerCount'), "Storage (MB)") -ForegroundColor $Global:ColorScheme.Header
            Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

            $rank = 1
            foreach ($sender in $topSenders) {
                $domainDisplay = if ($sender.Domain.Length -gt 44) {
                    $sender.Domain.Substring(0, 41) + "..."
                } else {
                    $sender.Domain
                }

                Write-Host ($format -f $rank, $domainDisplay, $sender.Count, $sender.TotalSizeMB) -ForegroundColor $Global:ColorScheme.Normal
                $rank++
            }
        }

        # Display Storage Analysis Section
        Write-Host "`n$(Get-LocalizedString 'analytics_sectionStorage')" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

        $storageData = Get-StorageByDomain -TopCount 10

        if ($storageData.Count -gt 0) {
            $format = "{0,-4} {1,-40} {2,12} {3,10} {4,12}"
            Write-Host ($format -f "#", (Get-LocalizedString 'senderOverview_headerDomain'),
                "Storage (MB)", (Get-LocalizedString 'senderOverview_headerCount'), "Avg Size (KB)") -ForegroundColor $Global:ColorScheme.Header
            Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

            $rank = 1
            foreach ($item in $storageData) {
                $domainDisplay = if ($item.Domain.Length -gt 39) {
                    $item.Domain.Substring(0, 36) + "..."
                } else {
                    $item.Domain
                }

                Write-Host ($format -f $rank, $domainDisplay, $item.TotalSizeMB, $item.EmailCount, $item.AvgSizeKB) -ForegroundColor $Global:ColorScheme.Normal
                $rank++
            }
        }

        # Display Growth Trend Section
        Write-Host "`n$(Get-LocalizedString 'analytics_sectionTrend')" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

        $trend = Get-MailboxGrowthTrend

        if ($trend.Count -gt 0) {
            # Show last 12 months or all if less
            $recentTrend = $trend | Select-Object -Last 12

            foreach ($period in $recentTrend) {
                $barLength = [math]::Min([math]::Round($period.EmailCount / 10), 50)
                $bar = "█" * $barLength

                Write-Host "  $($period.Period): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
                Write-Host "$bar " -NoNewline -ForegroundColor $Global:ColorScheme.Highlight
                Write-Host "($($period.EmailCount))" -ForegroundColor $Global:ColorScheme.Value
            }
        }

        # Display Unsubscribe Opportunities Section
        Write-Host "`n$(Get-LocalizedString 'analytics_sectionUnsubscribe')" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

        $opportunities = Get-UnsubscribeOpportunities -MinEmailCount 5

        if ($opportunities.Count -gt 0) {
            Write-Host "  $(Get-LocalizedString 'analytics_foundOpportunities' -FormatArgs @($opportunities.Count))" -ForegroundColor $Global:ColorScheme.Info
            Write-Host ""

            $topOpportunities = $opportunities | Select-Object -First 5

            foreach ($opp in $topOpportunities) {
                Write-Host "  • " -NoNewline -ForegroundColor $Global:ColorScheme.Highlight
                Write-Host "$($opp.Domain) " -NoNewline -ForegroundColor $Global:ColorScheme.Value
                Write-Host "($($opp.EmailCount) $(Get-LocalizedString 'analytics_emails'), $($opp.TotalSizeMB) MB)" -ForegroundColor $Global:ColorScheme.Normal
                Write-Host "    $($opp.Reasons)" -ForegroundColor $Global:ColorScheme.Muted
            }

            if ($opportunities.Count -gt 5) {
                Write-Host "`n  $(Get-LocalizedString 'analytics_moreOpportunities' -FormatArgs @($opportunities.Count - 5))" -ForegroundColor $Global:ColorScheme.Muted
            }
        } else {
            Write-Host "  $(Get-LocalizedString 'analytics_noOpportunities')" -ForegroundColor $Global:ColorScheme.Info
        }

        # Display recommendations
        Write-Host "`n$(Get-LocalizedString 'analytics_sectionRecommendations')" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

        # Generate recommendations based on data
        $recommendations = @()

        if ($stats.TotalSizeMB -gt 1000) {
            $recommendations += "• $(Get-LocalizedString 'analytics_recommendLargeSize')"
        }

        if ($opportunities.Count -gt 10) {
            $recommendations += "• $(Get-LocalizedString 'analytics_recommendNewsletters')"
        }

        if ($stats.EmailsWithAttachments -gt 100) {
            $recommendations += "• $(Get-LocalizedString 'analytics_recommendAttachments')"
        }

        $oldDomains = Get-TopSenders -TopCount 20 | Where-Object { $_.Count -gt 50 }
        if ($oldDomains.Count -gt 5) {
            $recommendations += "• $(Get-LocalizedString 'analytics_recommendHighVolume')"
        }

        if ($recommendations.Count -gt 0) {
            foreach ($rec in $recommendations) {
                Write-Host "  $rec" -ForegroundColor $Global:ColorScheme.Info
            }
        } else {
            Write-Host "  $(Get-LocalizedString 'analytics_noRecommendations')" -ForegroundColor $Global:ColorScheme.Success
        }

        Write-Host "`n" -NoNewline
        Write-Host ("=" * 80) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
    catch {
        Write-Error "Error displaying analytics dashboard: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Export functions
Export-ModuleMember -Function Get-MailboxStatistics, Get-TopSenders, Get-MailboxGrowthTrend, `
    Get-StorageByDomain, Get-UnsubscribeOpportunities, Show-AnalyticsDashboard
