<#
.SYNOPSIS
    Mailbox Health Monitor module for MailCleanBuddy
.DESCRIPTION
    Monitors mailbox health, provides warnings for quota, unusual patterns, and generates health scores.
#>

# Import dependencies

# Health data database path
$script:HealthDataPath = $null

# Function: Initialize-HealthMonitor
function Initialize-HealthMonitor {
    <#
    .SYNOPSIS
        Initializes health monitor database
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
    $script:HealthDataPath = Join-Path $PSScriptRoot "..\..\health_data_$sanitizedEmail.json"

    if (-not (Test-Path $script:HealthDataPath)) {
        $initialData = @{
            HealthSnapshots = @()
            Alerts = @()
            LastCheck = $null
        }
        $initialData | ConvertTo-Json -Depth 10 | Set-Content -Path $script:HealthDataPath -Encoding UTF8
    }
}

# Function: Get-MailboxHealth
function Get-MailboxHealth {
    <#
    .SYNOPSIS
        Analyzes current mailbox health
    .PARAMETER UserEmail
        User email address
    .OUTPUTS
        Health analysis object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Write-Host ""
        Write-Host (Get-LocalizedString "health_analyzing") -ForegroundColor $Global:ColorScheme.Info

        $cache = Get-SenderCache
        $health = [PSCustomObject]@{
            Score = 100
            TotalEmails = 0
            TotalSizeMB = 0
            UnreadEmails = 0
            UnreadPercentage = 0
            OldestEmailDays = 0
            LargestSenders = @()
            Warnings = @()
            Recommendations = @()
            Grade = "A"
            Timestamp = (Get-Date).ToString("o")
        }

        # Calculate statistics
        $allMessages = @()
        foreach ($domain in $cache.Keys) {
            $allMessages += $cache[$domain].Messages
        }

        $health.TotalEmails = $allMessages.Count

        $totalSize = 0
        $unreadCount = 0
        $oldestDate = [DateTime]::Now

        foreach ($msg in $allMessages) {
            if ($msg.Size) {
                $totalSize += $msg.Size
            }
            if (-not $msg.IsRead) {
                $unreadCount++
            }

            $receivedDate = ConvertTo-SafeDateTime -DateTimeValue $msg.ReceivedDateTime
            if ($receivedDate -lt $oldestDate) {
                $oldestDate = $receivedDate
            }
        }

        $health.TotalSizeMB = [math]::Round($totalSize / 1MB, 2)
        $health.UnreadEmails = $unreadCount
        if ($health.TotalEmails -gt 0) {
            $health.UnreadPercentage = [math]::Round(($unreadCount / $health.TotalEmails) * 100, 1)
        }
        $health.OldestEmailDays = ([DateTime]::Now - $oldestDate).Days

        # Find largest senders
        $senderSizes = @{}
        foreach ($domain in $cache.Keys) {
            $domainSize = 0
            foreach ($msg in $cache[$domain].Messages) {
                $domainSize += $msg.Size
            }
            $senderSizes[$domain] = $domainSize
        }

        $health.LargestSenders = $senderSizes.GetEnumerator() |
            Sort-Object Value -Descending |
            Select-Object -First 5 |
            ForEach-Object {
                [PSCustomObject]@{
                    Domain = $_.Key
                    SizeMB = [math]::Round($_.Value / 1MB, 2)
                }
            }

        # Analyze and generate warnings
        $deductions = 0

        # Check 1: Storage size
        if ($health.TotalSizeMB -gt 5000) {
            $health.Warnings += (Get-LocalizedString "health_warnLargeMailbox")
            $health.Recommendations += (Get-LocalizedString "health_recArchiveOld")
            $deductions += 15
        } elseif ($health.TotalSizeMB -gt 2000) {
            $health.Warnings += (Get-LocalizedString "health_warnMediumMailbox")
            $health.Recommendations += (Get-LocalizedString "health_recConsiderArchiving")
            $deductions += 10
        }

        # Check 2: Unread percentage
        if ($health.UnreadPercentage -gt 50) {
            $health.Warnings += (Get-LocalizedString "health_warnHighUnread")
            $health.Recommendations += (Get-LocalizedString "health_recProcessUnread")
            $deductions += 10
        } elseif ($health.UnreadPercentage -gt 25) {
            $health.Warnings += (Get-LocalizedString "health_warnMediumUnread")
            $health.Recommendations += (Get-LocalizedString "health_recReviewUnread")
            $deductions += 5
        }

        # Check 3: Old emails
        if ($health.OldestEmailDays -gt 730) { # 2 years
            $health.Warnings += (Get-LocalizedString "health_warnVeryOldEmails")
            $health.Recommendations += (Get-LocalizedString "health_recArchiveAncient")
            $deductions += 10
        } elseif ($health.OldestEmailDays -gt 365) { # 1 year
            $health.Warnings += (Get-LocalizedString "health_warnOldEmails")
            $health.Recommendations += (Get-LocalizedString "health_recReviewOld")
            $deductions += 5
        }

        # Check 4: Total email count
        if ($health.TotalEmails -gt 50000) {
            $health.Warnings += (Get-LocalizedString "health_warnTooManyEmails")
            $health.Recommendations += (Get-LocalizedString "health_recBulkClean")
            $deductions += 15
        } elseif ($health.TotalEmails -gt 20000) {
            $health.Warnings += (Get-LocalizedString "health_warnManyEmails")
            $health.Recommendations += (Get-LocalizedString "health_recRegularClean")
            $deductions += 10
        }

        # Check 5: Sender concentration
        if ($health.LargestSenders.Count -gt 0) {
            $topSenderPercentage = ($cache[$health.LargestSenders[0].Domain].Messages.Count / $health.TotalEmails) * 100
            if ($topSenderPercentage -gt 30) {
                $health.Warnings += (Get-LocalizedString "health_warnSenderConcentration" -FormatArgs @($health.LargestSenders[0].Domain))
                $health.Recommendations += (Get-LocalizedString "health_recUnsubscribeOrFilter")
                $deductions += 5
            }
        }

        # Calculate final score
        $health.Score = [math]::Max(0, 100 - $deductions)

        # Assign grade
        if ($health.Score -ge 90) {
            $health.Grade = "A"
        } elseif ($health.Score -ge 80) {
            $health.Grade = "B"
        } elseif ($health.Score -ge 70) {
            $health.Grade = "C"
        } elseif ($health.Score -ge 60) {
            $health.Grade = "D"
        } else {
            $health.Grade = "F"
        }

        # Add positive feedback if healthy
        if ($health.Warnings.Count -eq 0) {
            $health.Recommendations += (Get-LocalizedString "health_recKeepItUp")
        }

        return $health
    }
    catch {
        Write-Error "Error analyzing mailbox health: $($_.Exception.Message)"
        return $null
    }
}

# Function: Show-HealthMonitor
function Show-HealthMonitor {
    <#
    .SYNOPSIS
        Interactive health monitor interface
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Initialize-HealthMonitor -UserEmail $UserEmail

        Clear-Host

        $title = Get-LocalizedString "health_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Get current health
        $health = Get-MailboxHealth -UserEmail $UserEmail

        if (-not $health) {
            Write-Host (Get-LocalizedString "health_errorAnalyzing") -ForegroundColor $Global:ColorScheme.Error
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Save snapshot
        Save-HealthSnapshot -Health $health

        # Display health score
        Write-Host ""
        Write-Host (Get-LocalizedString "health_scoreTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        $scoreColor = if ($health.Score -ge 80) {
            $Global:ColorScheme.Success
        } elseif ($health.Score -ge 60) {
            $Global:ColorScheme.Warning
        } else {
            $Global:ColorScheme.Error
        }

        Write-Host "  $(Get-LocalizedString 'health_overallScore'): " -NoNewline
        Write-Host "$($health.Score)/100 (Grade: $($health.Grade))" -ForegroundColor $scoreColor
        Write-Host ""

        # Display statistics
        Write-Host (Get-LocalizedString "health_statisticsTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""
        Write-Host "  $(Get-LocalizedString 'analytics_totalEmails'): " -NoNewline
        Write-Host "$($health.TotalEmails)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "  $(Get-LocalizedString 'analytics_totalStorage'): " -NoNewline
        Write-Host "$($health.TotalSizeMB) MB" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "  $(Get-LocalizedString 'health_unreadEmails'): " -NoNewline
        Write-Host "$($health.UnreadEmails) ($($health.UnreadPercentage)%)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "  $(Get-LocalizedString 'health_oldestEmail'): " -NoNewline
        Write-Host "$($health.OldestEmailDays) $(Get-LocalizedString 'health_daysAgo')" -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""

        # Display warnings
        if ($health.Warnings.Count -gt 0) {
            Write-Host (Get-LocalizedString "health_warningsTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""
            foreach ($warning in $health.Warnings) {
                Write-Host "  ⚠️  $warning" -ForegroundColor $Global:ColorScheme.Warning
            }
            Write-Host ""
        }

        # Display recommendations
        if ($health.Recommendations.Count -gt 0) {
            Write-Host (Get-LocalizedString "health_recommendationsTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""
            foreach ($rec in $health.Recommendations) {
                Write-Host "  💡 $rec" -ForegroundColor $Global:ColorScheme.Info
            }
            Write-Host ""
        }

        # Display top storage consumers
        if ($health.LargestSenders.Count -gt 0) {
            Write-Host (Get-LocalizedString "health_topStorageTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""
            $format = "  {0,-3} {1,-50} {2,15}"
            Write-Host ($format -f "#", "Domain", "Storage (MB)") -ForegroundColor $Global:ColorScheme.Header
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

            $index = 1
            foreach ($sender in $health.LargestSenders) {
                Write-Host ($format -f $index, $sender.Domain, $sender.SizeMB) -ForegroundColor $Global:ColorScheme.Normal
                $index++
            }
            Write-Host ""
        }

        # Menu
        Write-Host (Get-LocalizedString "health_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  1. $(Get-LocalizedString 'health_viewHistory')" -ForegroundColor Cyan
        Write-Host "  2. $(Get-LocalizedString 'health_exportReport')" -ForegroundColor Magenta
        Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
        Write-Host ""

        $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

        switch ($choice.ToUpper()) {
            "1" {
                Show-HealthHistory
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            }
            "2" {
                Export-HealthReport -Health $health
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            }
        }
    }
    catch {
        Write-Error "Error in health monitor: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Save-HealthSnapshot
function Save-HealthSnapshot {
    <#
    .SYNOPSIS
        Saves health snapshot to history
    .PARAMETER Health
        Health analysis object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Health
    )

    try {
        $data = Get-Content -Path $script:HealthDataPath -Raw | ConvertFrom-Json

        $snapshot = [PSCustomObject]@{
            Score = $Health.Score
            Grade = $Health.Grade
            TotalEmails = $Health.TotalEmails
            TotalSizeMB = $Health.TotalSizeMB
            UnreadEmails = $Health.UnreadEmails
            UnreadPercentage = $Health.UnreadPercentage
            WarningCount = $Health.Warnings.Count
            Timestamp = $Health.Timestamp
        }

        $data.HealthSnapshots += $snapshot
        $data.LastCheck = $Health.Timestamp

        # Keep only last 30 snapshots
        if ($data.HealthSnapshots.Count -gt 30) {
            $data.HealthSnapshots = $data.HealthSnapshots | Select-Object -Last 30
        }

        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:HealthDataPath -Encoding UTF8
    }
    catch {
        Write-Warning "Could not save health snapshot: $($_.Exception.Message)"
    }
}

# Function: Show-HealthHistory
function Show-HealthHistory {
    <#
    .SYNOPSIS
        Shows health history and trends
    #>
    [CmdletBinding()]
    param()

    $data = Get-Content -Path $script:HealthDataPath -Raw | ConvertFrom-Json

    if ($data.HealthSnapshots.Count -eq 0) {
        Write-Host "`n$(Get-LocalizedString 'health_noHistory')" -ForegroundColor $Global:ColorScheme.Warning
        return
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "health_historyTitle" -FormatArgs @($data.HealthSnapshots.Count)) -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

    $format = "  {0,-20} {1,10} {2,8} {3,12} {4,12} {5,10}"
    Write-Host ($format -f "Date", "Score", "Grade", "Emails", "Size (MB)", "Warnings") -ForegroundColor $Global:ColorScheme.Header
    Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

    $recent = $data.HealthSnapshots | Select-Object -Last 10 | Sort-Object { ConvertTo-SafeDateTime -DateTimeValue $_.Timestamp } -Descending

    foreach ($snapshot in $recent) {
        $timestamp = ConvertTo-SafeDateTime -DateTimeValue $snapshot.Timestamp.ToString('yyyy-MM-dd HH:mm')

        $scoreColor = if ($snapshot.Score -ge 80) {
            $Global:ColorScheme.Success
        } elseif ($snapshot.Score -ge 60) {
            $Global:ColorScheme.Warning
        } else {
            $Global:ColorScheme.Error
        }

        Write-Host ($format -f $timestamp, $snapshot.Score, $snapshot.Grade, $snapshot.TotalEmails, $snapshot.TotalSizeMB, $snapshot.WarningCount) `
            -ForegroundColor $scoreColor
    }

    Write-Host ""

    # Calculate trend
    if ($data.HealthSnapshots.Count -ge 2) {
        $current = $data.HealthSnapshots[-1]
        $previous = $data.HealthSnapshots[-2]

        $scoreTrend = $current.Score - $previous.Score
        $emailTrend = $current.TotalEmails - $previous.TotalEmails
        $sizeTrend = $current.TotalSizeMB - $previous.TotalSizeMB

        Write-Host (Get-LocalizedString "health_trendsTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ""

        $trendSymbol = if ($scoreTrend -gt 0) { "↑" } elseif ($scoreTrend -lt 0) { "↓" } else { "→" }
        $trendColor = if ($scoreTrend -gt 0) { $Global:ColorScheme.Success } elseif ($scoreTrend -lt 0) { $Global:ColorScheme.Warning } else { $Global:ColorScheme.Normal }
        Write-Host "  $(Get-LocalizedString 'health_scoreChange'): " -NoNewline
        Write-Host "$trendSymbol $scoreTrend points" -ForegroundColor $trendColor

        $emailSymbol = if ($emailTrend -lt 0) { "↓" } elseif ($emailTrend -gt 0) { "↑" } else { "→" }
        $emailColor = if ($emailTrend -lt 0) { $Global:ColorScheme.Success } elseif ($emailTrend -gt 100) { $Global:ColorScheme.Warning } else { $Global:ColorScheme.Normal }
        Write-Host "  $(Get-LocalizedString 'health_emailCountChange'): " -NoNewline
        Write-Host "$emailSymbol $emailTrend emails" -ForegroundColor $emailColor

        $sizeSymbol = if ($sizeTrend -lt 0) { "↓" } elseif ($sizeTrend -gt 0) { "↑" } else { "→" }
        $sizeColor = if ($sizeTrend -lt 0) { $Global:ColorScheme.Success } elseif ($sizeTrend -gt 100) { $Global:ColorScheme.Warning } else { $Global:ColorScheme.Normal }
        Write-Host "  $(Get-LocalizedString 'health_sizeChange'): " -NoNewline
        Write-Host "$sizeSymbol $([math]::Round($sizeTrend, 2)) MB" -ForegroundColor $sizeColor
        Write-Host ""
    }
}

# Function: Export-HealthReport
function Export-HealthReport {
    <#
    .SYNOPSIS
        Exports health report to CSV
    .PARAMETER Health
        Health analysis object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Health
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\health_report.csv"
    $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    $reportData = [PSCustomObject]@{
        Timestamp = $Health.Timestamp
        Score = $Health.Score
        Grade = $Health.Grade
        TotalEmails = $Health.TotalEmails
        TotalSizeMB = $Health.TotalSizeMB
        UnreadEmails = $Health.UnreadEmails
        UnreadPercentage = $Health.UnreadPercentage
        OldestEmailDays = $Health.OldestEmailDays
        WarningCount = $Health.Warnings.Count
        RecommendationCount = $Health.Recommendations.Count
    }

    $reportData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
}

# Export functions
Export-ModuleMember -Function Show-HealthMonitor, Get-MailboxHealth, Initialize-HealthMonitor
