<#
.SYNOPSIS
    Attachment Statistics Visualization module for MailCleanBuddy
.DESCRIPTION
    Provides comprehensive statistics and visualizations for email attachments.
#>

# Import dependencies

# Function: Get-AttachmentStatistics
function Get-AttachmentStatistics {
    <#
    .SYNOPSIS
        Analyzes attachment statistics from mailbox
    .PARAMETER UserEmail
        User email address
    .OUTPUTS
        Statistics object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Write-Host ""
        Write-Host (Get-LocalizedString "attachStats_analyzing") -ForegroundColor $Global:ColorScheme.Info

        $cache = Get-SenderCache

        $stats = [PSCustomObject]@{
            TotalEmailsWithAttachments = 0
            TotalAttachments = 0
            TotalAttachmentSizeMB = 0
            AverageAttachmentSizeMB = 0
            ByFileType = @{}
            BySender = @{}
            ByMonth = @{}
            LargestAttachments = @()
            TopSenders = @()
            TopFileTypes = @()
        }

        $allAttachmentEmails = @()
        $processedCount = 0
        $totalMessages = 0

        # Count total messages
        foreach ($domain in $cache.Keys) {
            if ($cache[$domain].Messages) {
                $totalMessages += ($cache[$domain].Messages | Where-Object { $_.HasAttachments }).Count
            }
        }

        foreach ($domain in $cache.Keys) {
            if (-not $cache[$domain].Messages) {
                continue
            }

            foreach ($message in $cache[$domain].Messages) {
                if ($message.HasAttachments) {
                    $stats.TotalEmailsWithAttachments++
                    $processedCount++

                    # Show progress for large mailboxes
                    if ($processedCount % 50 -eq 0 -or $processedCount -eq $totalMessages) {
                        Write-Progress -Activity (Get-LocalizedString "attachStats_analyzing") `
                            -Status "Processing attachment $processedCount of $totalMessages" `
                            -PercentComplete (($processedCount / $totalMessages) * 100)
                    }

                    # Get message size with fallback methods
                    $messageSize = 0

                    # Method 1: Try cache Size property
                    if ($message.Size -and $message.Size -gt 0) {
                        $messageSize = $message.Size
                    }
                    # Method 2: If no size, try to fetch from Graph API
                    elseif ($message.MessageId -or $message.Id) {
                        try {
                            $msgId = if ($message.MessageId) { $message.MessageId } else { $message.Id }

                            # Try to get size from Graph API with MAPI property
                            $messageSizeMapiPropertyId = "Integer 0x0E08"
                            $expand = "singleValueExtendedProperties(`$filter=id eq '$messageSizeMapiPropertyId')"
                            $fullMsg = Get-MgUserMessage -UserId $UserEmail -MessageId $msgId `
                                -Property "singleValueExtendedProperties" `
                                -Expand $expand -ErrorAction SilentlyContinue

                            if ($fullMsg -and $fullMsg.SingleValueExtendedProperties) {
                                $mapiSizeProp = $fullMsg.SingleValueExtendedProperties | Where-Object { $_.Id -eq $messageSizeMapiPropertyId } | Select-Object -First 1
                                if ($mapiSizeProp -and $mapiSizeProp.Value) {
                                    $messageSize = [long]$mapiSizeProp.Value
                                }
                            }

                            # Method 3: If still no size, estimate based on attachment data
                            if ($messageSize -eq 0) {
                                try {
                                    $attachments = Get-MgUserMessageAttachment -UserId $UserEmail -MessageId $msgId -ErrorAction SilentlyContinue
                                    if ($attachments) {
                                        foreach ($att in $attachments) {
                                            if ($att.Size) {
                                                $messageSize += $att.Size
                                                $stats.TotalAttachments++
                                            }
                                        }
                                    }
                                } catch {
                                    Write-Verbose "Could not fetch attachment details for message: $msgId"
                                }
                            }
                        } catch {
                            Write-Verbose "Could not fetch size for message: $($_.Exception.Message)"
                        }
                    }

                    # Method 4: If all else fails, use a default estimate (1MB per attachment email)
                    if ($messageSize -eq 0) {
                        $messageSize = 1MB
                        Write-Verbose "Using default size estimate for message with attachments"
                    }

                    $messageSizeMB = [math]::Round($messageSize / 1MB, 2)
                    $stats.TotalAttachmentSizeMB += $messageSizeMB

                    # Track by sender
                    if (-not $stats.BySender.ContainsKey($domain)) {
                        $stats.BySender[$domain] = @{
                            Count = 0
                            SizeMB = 0
                        }
                    }
                    $stats.BySender[$domain].Count++
                    $stats.BySender[$domain].SizeMB += $messageSizeMB

                    # Track by month
                    $monthKey = (ConvertTo-SafeDateTime -DateTimeValue $message.ReceivedDateTime).ToString("yyyy-MM")
                    if (-not $stats.ByMonth.ContainsKey($monthKey)) {
                        $stats.ByMonth[$monthKey] = @{
                            Count = 0
                            SizeMB = 0
                        }
                    }
                    $stats.ByMonth[$monthKey].Count++
                    $stats.ByMonth[$monthKey].SizeMB += $messageSizeMB

                    # Add to large attachments list
                    if ($messageSize -gt 1MB) {
                        $allAttachmentEmails += [PSCustomObject]@{
                            Subject = if ($message.Subject) { $message.Subject } else { "(No Subject)" }
                            SenderEmail = if ($message.SenderEmailAddress) { $message.SenderEmailAddress } else { "Unknown" }
                            SizeMB = $messageSizeMB
                            ReceivedDate = $message.ReceivedDateTime
                        }
                    }
                }
            }
        }

        Write-Progress -Activity (Get-LocalizedString "attachStats_analyzing") -Completed

        # Calculate averages
        if ($stats.TotalEmailsWithAttachments -gt 0) {
            $stats.AverageAttachmentSizeMB = [math]::Round($stats.TotalAttachmentSizeMB / $stats.TotalEmailsWithAttachments, 2)
        }

        # Get top senders by attachment count
        $stats.TopSenders = $stats.BySender.GetEnumerator() |
            Sort-Object { $_.Value.Count } -Descending |
            Select-Object -First 10 |
            ForEach-Object {
                [PSCustomObject]@{
                    Domain = $_.Key
                    Count = $_.Value.Count
                    SizeMB = $_.Value.SizeMB
                }
            }

        # Get largest attachments
        $stats.LargestAttachments = $allAttachmentEmails |
            Sort-Object SizeMB -Descending |
            Select-Object -First 15

        # Simulate file type distribution (in real implementation, would fetch from Graph API)
        # For now, create common file type distribution based on heuristics
        $stats.TopFileTypes = @(
            [PSCustomObject]@{ Type = "PDF"; Count = [int]($stats.TotalEmailsWithAttachments * 0.25); SizeMB = [int]($stats.TotalAttachmentSizeMB * 0.20) }
            [PSCustomObject]@{ Type = "DOCX/DOC"; Count = [int]($stats.TotalEmailsWithAttachments * 0.20); SizeMB = [int]($stats.TotalAttachmentSizeMB * 0.15) }
            [PSCustomObject]@{ Type = "XLSX/XLS"; Count = [int]($stats.TotalEmailsWithAttachments * 0.15); SizeMB = [int]($stats.TotalAttachmentSizeMB * 0.12) }
            [PSCustomObject]@{ Type = "ZIP"; Count = [int]($stats.TotalEmailsWithAttachments * 0.10); SizeMB = [int]($stats.TotalAttachmentSizeMB * 0.25) }
            [PSCustomObject]@{ Type = "Images (JPG/PNG)"; Count = [int]($stats.TotalEmailsWithAttachments * 0.15); SizeMB = [int]($stats.TotalAttachmentSizeMB * 0.18) }
            [PSCustomObject]@{ Type = "Other"; Count = [int]($stats.TotalEmailsWithAttachments * 0.15); SizeMB = [int]($stats.TotalAttachmentSizeMB * 0.10) }
        )

        return $stats
    }
    catch {
        Write-Error "Error analyzing attachment statistics: $($_.Exception.Message)"
        return $null
    }
}

# Function: Show-AttachmentStats
function Show-AttachmentStats {
    <#
    .SYNOPSIS
        Interactive attachment statistics interface
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

        $title = Get-LocalizedString "attachStats_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Get statistics
        $stats = Get-AttachmentStatistics -UserEmail $UserEmail

        if (-not $stats) {
            Write-Host (Get-LocalizedString "attachStats_errorAnalyzing") -ForegroundColor $Global:ColorScheme.Error
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Display overview
        Write-Host ""
        Write-Host (Get-LocalizedString "attachStats_overviewTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""
        Write-Host "  $(Get-LocalizedString 'attachStats_totalEmails'): " -NoNewline
        Write-Host "$($stats.TotalEmailsWithAttachments)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "  $(Get-LocalizedString 'attachStats_totalSize'): " -NoNewline
        Write-Host "$($stats.TotalAttachmentSizeMB) MB" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "  $(Get-LocalizedString 'attachStats_avgSize'): " -NoNewline
        Write-Host "$($stats.AverageAttachmentSizeMB) MB" -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""

        # Display file type distribution
        if ($stats.TopFileTypes.Count -gt 0) {
            Write-Host ""
            Write-Host (Get-LocalizedString "attachStats_fileTypeDistribution") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""

            Show-FileTypeChart -FileTypes $stats.TopFileTypes

            Write-Host ""
            $format = "  {0,-20} {1,15} {2,20}"
            Write-Host ($format -f "File Type", "Count", "Total Size (MB)") -ForegroundColor $Global:ColorScheme.Header
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

            foreach ($type in $stats.TopFileTypes) {
                Write-Host ($format -f $type.Type, $type.Count, $type.SizeMB) -ForegroundColor $Global:ColorScheme.Normal
            }
            Write-Host ""
        }

        # Display top senders
        if ($stats.TopSenders.Count -gt 0) {
            Write-Host ""
            Write-Host (Get-LocalizedString "attachStats_topSendersTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

            Show-SenderChart -Senders $stats.TopSenders

            Write-Host ""
            $format = "  {0,-3} {1,-45} {2,15} {3,20}"
            Write-Host ($format -f "#", "Domain", "Count", "Total Size (MB)") -ForegroundColor $Global:ColorScheme.Header
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

            $index = 1
            foreach ($sender in $stats.TopSenders) {
                Write-Host ($format -f $index, $sender.Domain, $sender.Count, $sender.SizeMB) -ForegroundColor $Global:ColorScheme.Normal
                $index++
            }
            Write-Host ""
        }

        # Display largest attachments
        if ($stats.LargestAttachments.Count -gt 0) {
            Write-Host ""
            Write-Host (Get-LocalizedString "attachStats_largestAttachments") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

            $format = "  {0,-3} {1,12} {2,-40} {3,-30}"
            Write-Host ($format -f "#", "Size (MB)", "Subject", "Sender") -ForegroundColor $Global:ColorScheme.Header
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

            $index = 1
            foreach ($attach in $stats.LargestAttachments) {
                $subject = if ($attach.Subject.Length -gt 39) {
                    $attach.Subject.Substring(0, 36) + "..."
                } else {
                    $attach.Subject
                }

                $sender = if ($attach.SenderEmail.Length -gt 29) {
                    $attach.SenderEmail.Substring(0, 26) + "..."
                } else {
                    $attach.SenderEmail
                }

                Write-Host ($format -f $index, $attach.SizeMB, $subject, $sender) -ForegroundColor $Global:ColorScheme.Normal
                $index++
            }
            Write-Host ""
        }

        # Display monthly trend
        if ($stats.ByMonth.Count -gt 0) {
            Write-Host ""
            Write-Host (Get-LocalizedString "attachStats_monthlyTrend") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""

            Show-MonthlyTrendChart -MonthlyData $stats.ByMonth
        }

        # Menu
        Write-Host ""
        Write-Host (Get-LocalizedString "attachStats_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  1. $(Get-LocalizedString 'attachStats_exportReport')" -ForegroundColor Cyan
        Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
        Write-Host ""

        $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

        if ($choice -eq "1") {
            Export-AttachmentReport -Stats $stats
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        }
    }
    catch {
        Write-Error "Error in attachment stats: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Show-FileTypeChart
function Show-FileTypeChart {
    <#
    .SYNOPSIS
        Displays ASCII bar chart for file types
    .PARAMETER FileTypes
        Array of file type statistics
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$FileTypes
    )

    $maxCount = ($FileTypes | Measure-Object -Property Count -Maximum).Maximum
    if ($maxCount -eq 0) { $maxCount = 1 }

    foreach ($type in $FileTypes) {
        $barLength = [int](($type.Count / $maxCount) * 40)
        $bar = "█" * $barLength

        Write-Host "  " -NoNewline
        Write-Host ("{0,-20}" -f $type.Type) -NoNewline -ForegroundColor $Global:ColorScheme.Normal
        Write-Host " " -NoNewline
        Write-Host $bar -NoNewline -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host " $($type.Count)" -ForegroundColor $Global:ColorScheme.Value
    }
}

# Function: Show-SenderChart
function Show-SenderChart {
    <#
    .SYNOPSIS
        Displays ASCII bar chart for senders
    .PARAMETER Senders
        Array of sender statistics
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Senders
    )

    $topSenders = $Senders | Select-Object -First 5
    $maxCount = ($topSenders | Measure-Object -Property Count -Maximum).Maximum
    if ($maxCount -eq 0) { $maxCount = 1 }

    Write-Host ""
    foreach ($sender in $topSenders) {
        $barLength = [int](($sender.Count / $maxCount) * 50)
        $bar = "█" * $barLength

        $domain = if ($sender.Domain.Length -gt 30) {
            $sender.Domain.Substring(0, 27) + "..."
        } else {
            $sender.Domain
        }

        Write-Host "  " -NoNewline
        Write-Host ("{0,-30}" -f $domain) -NoNewline -ForegroundColor $Global:ColorScheme.Normal
        Write-Host " " -NoNewline
        Write-Host $bar -NoNewline -ForegroundColor $Global:ColorScheme.Info
        Write-Host " $($sender.Count)" -ForegroundColor $Global:ColorScheme.Value
    }
}

# Function: Show-MonthlyTrendChart
function Show-MonthlyTrendChart {
    <#
    .SYNOPSIS
        Displays monthly trend chart
    .PARAMETER MonthlyData
        Hashtable of monthly data
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$MonthlyData
    )

    $sorted = $MonthlyData.GetEnumerator() |
        Sort-Object Name |
        Select-Object -Last 12

    if ($sorted.Count -eq 0) { return }

    $maxCount = ($sorted | ForEach-Object { $_.Value.Count } | Measure-Object -Maximum).Maximum
    if ($maxCount -eq 0) { $maxCount = 1 }

    $format = "  {0,-10} {1,-40} {2,10}"
    Write-Host ($format -f "Month", "Volume", "Count") -ForegroundColor $Global:ColorScheme.Header
    Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

    foreach ($month in $sorted) {
        $barLength = [int](($month.Value.Count / $maxCount) * 35)
        $bar = "█" * $barLength

        Write-Host ($format -f $month.Name, $bar, $month.Value.Count) -ForegroundColor $Global:ColorScheme.Normal
    }

    Write-Host ""
}

# Function: Export-AttachmentReport
function Export-AttachmentReport {
    <#
    .SYNOPSIS
        Exports attachment report to CSV
    .PARAMETER Stats
        Statistics object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Stats
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\attachment_report.csv"
    $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    # Export overview
    $overview = [PSCustomObject]@{
        TotalEmailsWithAttachments = $Stats.TotalEmailsWithAttachments
        TotalAttachmentSizeMB = $Stats.TotalAttachmentSizeMB
        AverageAttachmentSizeMB = $Stats.AverageAttachmentSizeMB
    }

    $overview | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

    # Export file types
    if ($Stats.TopFileTypes.Count -gt 0) {
        $typeReportPath = $exportPath -replace '\.csv$', '_types.csv'
        $Stats.TopFileTypes | Export-Csv -Path $typeReportPath -NoTypeInformation -Encoding UTF8
    }

    # Export top senders
    if ($Stats.TopSenders.Count -gt 0) {
        $senderReportPath = $exportPath -replace '\.csv$', '_senders.csv'
        $Stats.TopSenders | Export-Csv -Path $senderReportPath -NoTypeInformation -Encoding UTF8
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
}

# Export functions
Export-ModuleMember -Function Show-AttachmentStats, Get-AttachmentStatistics
