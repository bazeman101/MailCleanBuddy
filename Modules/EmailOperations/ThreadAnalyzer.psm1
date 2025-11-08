<#
.SYNOPSIS
    Thread Analyzer module for MailCleanBuddy
.DESCRIPTION
    Analyzes email conversations and threads for bulk management.
    Groups emails by conversation ID and provides thread statistics.
#>

# Import dependencies

# Function: Get-EmailThreads
function Get-EmailThreads {
    <#
    .SYNOPSIS
        Groups emails by conversation/thread
    .OUTPUTS
        Hashtable with threads
    #>
    [CmdletBinding()]
    param()

    try {
        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            return @{}
        }

        $threads = @{}
        $allMessages = @()

        # Collect all messages
        foreach ($domain in $cache.Keys) {
            $allMessages += $cache[$domain].Messages
        }

        # Group by subject (simplified threading)
        # In real implementation, would use ConversationId from Graph API
        foreach ($msg in $allMessages) {
            # Normalize subject
            $normalizedSubject = if ($msg.Subject) {
                $msg.Subject -replace '^\s*(RE|FW|FWD|AW):\s*', '' -replace '\s+', ' '
            } else {
                "[No Subject]"
            }

            if (-not $threads.ContainsKey($normalizedSubject)) {
                $threads[$normalizedSubject] = @{
                    Subject = $normalizedSubject
                    Messages = @()
                    Participants = @()
                    TotalSize = 0
                }
            }

            $threads[$normalizedSubject].Messages += $msg
            $threads[$normalizedSubject].TotalSize += $msg.Size

            # Track unique participants
            if ($msg.SenderEmailAddress -and
                $threads[$normalizedSubject].Participants -notcontains $msg.SenderEmailAddress) {
                $threads[$normalizedSubject].Participants += $msg.SenderEmailAddress
            }
        }

        return $threads
    }
    catch {
        Write-Error "Error analyzing threads: $($_.Exception.Message)"
        return @{}
    }
}

# Function: Show-ThreadAnalyzer
function Show-ThreadAnalyzer {
    <#
    .SYNOPSIS
        Interactive thread analyzer interface
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

        # Display header
        $title = Get-LocalizedString "thread_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host (Get-LocalizedString "thread_analyzing") -ForegroundColor $Global:ColorScheme.Info

        # Analyze threads
        $threads = Get-EmailThreads

        if ($threads.Count -eq 0) {
            Write-Host "`n$(Get-LocalizedString 'thread_noThreads')" -ForegroundColor $Global:ColorScheme.Warning
            Write-Host ""
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Calculate statistics
        $totalThreads = $threads.Count
        $totalMessages = 0
        $threads.Values | ForEach-Object { $totalMessages += $_.Messages.Count }
        $avgMessagesPerThread = [math]::Round($totalMessages / $totalThreads, 1)

        # Display statistics
        Write-Host "`n$(Get-LocalizedString 'thread_foundThreads' -FormatArgs @($totalThreads))" -ForegroundColor $Global:ColorScheme.Success
        Write-Host "$(Get-LocalizedString 'thread_totalMessages' -FormatArgs @($totalMessages))" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "$(Get-LocalizedString 'thread_avgMessages' -FormatArgs @($avgMessagesPerThread))" -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Sort threads by message count
        $sortedThreads = $threads.GetEnumerator() | Sort-Object { $_.Value.Messages.Count } -Descending

        # Display top threads
        Write-Host (Get-LocalizedString "thread_topThreads") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        $format = "{0,-4} {1,-50} {2,10} {3,12} {4,10}"
        Write-Host ($format -f "#", (Get-LocalizedString 'standardizedList_headerSubject'),
            "Messages", "Size (MB)", "People") -ForegroundColor $Global:ColorScheme.Header
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        $topThreads = $sortedThreads | Select-Object -First 15
        $index = 1

        foreach ($threadEntry in $topThreads) {
            $thread = $threadEntry.Value

            $subject = if ($thread.Subject.Length -gt 49) {
                $thread.Subject.Substring(0, 46) + "..."
            } else {
                $thread.Subject
            }

            $sizeMB = [math]::Round($thread.TotalSize / 1MB, 2)

            Write-Host ($format -f $index, $subject, $thread.Messages.Count, $sizeMB, $thread.Participants.Count) `
                -ForegroundColor $Global:ColorScheme.Normal
            $index++
        }

        Write-Host ""

        # Menu for actions
        while ($true) {
            Write-Host (Get-LocalizedString "thread_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  1. $(Get-LocalizedString 'thread_viewThread')" -ForegroundColor Green
            Write-Host "  2. $(Get-LocalizedString 'thread_deleteThread')" -ForegroundColor Yellow
            Write-Host "  3. $(Get-LocalizedString 'thread_archiveThread')" -ForegroundColor Cyan
            Write-Host "  4. $(Get-LocalizedString 'thread_exportReport')" -ForegroundColor Magenta
            Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
            Write-Host ""

            $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

            switch ($choice.ToUpper()) {
                "1" {
                    # View thread
                    $threadNum = Read-Host (Get-LocalizedString "thread_enterNumber")
                    if ($threadNum -match '^\d+$' -and [int]$threadNum -ge 1 -and [int]$threadNum -le $topThreads.Count) {
                        $selectedThread = $topThreads[[int]$threadNum - 1].Value
                        Show-ThreadDetails -UserEmail $UserEmail -Thread $selectedThread
                    } else {
                        Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
                    }
                }
                "2" {
                    # Delete thread
                    $threadNum = Read-Host (Get-LocalizedString "thread_enterNumber")
                    if ($threadNum -match '^\d+$' -and [int]$threadNum -ge 1 -and [int]$threadNum -le $topThreads.Count) {
                        $selectedThread = $topThreads[[int]$threadNum - 1].Value
                        Invoke-DeleteThread -UserEmail $UserEmail -Thread $selectedThread
                    } else {
                        Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
                    }
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "3" {
                    # Archive thread
                    $threadNum = Read-Host (Get-LocalizedString "thread_enterNumber")
                    if ($threadNum -match '^\d+$' -and [int]$threadNum -ge 1 -and [int]$threadNum -le $topThreads.Count) {
                        $selectedThread = $topThreads[[int]$threadNum - 1].Value
                        Invoke-ArchiveThread -UserEmail $UserEmail -Thread $selectedThread
                    } else {
                        Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
                    }
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "4" {
                    # Export report
                    Export-ThreadReport -Threads $threads
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
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
        Write-Error "Error in thread analyzer: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Show-ThreadDetails
function Show-ThreadDetails {
    <#
    .SYNOPSIS
        Shows details of a thread
    .PARAMETER UserEmail
        User email address
    .PARAMETER Thread
        Thread object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [hashtable]$Thread
    )

    Clear-Host
    Write-Host "`n$(Get-LocalizedString 'thread_detailsTitle')" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    Write-Host "  $(Get-LocalizedString 'standardizedList_headerSubject'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $Thread.Subject -ForegroundColor $Global:ColorScheme.Value

    Write-Host "  Messages: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $Thread.Messages.Count -ForegroundColor $Global:ColorScheme.Value

    Write-Host "  Participants: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $Thread.Participants.Count -ForegroundColor $Global:ColorScheme.Value

    $sizeMB = [math]::Round($Thread.TotalSize / 1MB, 2)
    Write-Host "  Total Size: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host "$sizeMB MB" -ForegroundColor $Global:ColorScheme.Value

    Write-Host ""
    Write-Host (Get-LocalizedString "thread_messagesInThread") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

    $sortedMessages = $Thread.Messages | Sort-Object { ConvertTo-SafeDateTime -DateTimeValue $_.ReceivedDateTime }

    $format = "{0,-4} {1,-25} {2,-35} {3,20}"
    Write-Host ($format -f "#", (Get-LocalizedString 'standardizedList_headerSenderName'),
        "Email", (Get-LocalizedString 'standardizedList_headerDate')) -ForegroundColor $Global:ColorScheme.Header
    Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

    $index = 1
    foreach ($msg in $sortedMessages) {
        $sender = if ($msg.SenderName -and $msg.SenderName.Length -gt 24) {
            $msg.SenderName.Substring(0, 21) + "..."
        } elseif ($msg.SenderName) {
            $msg.SenderName
        } else {
            "Unknown"
        }

        $email = if ($msg.SenderEmailAddress -and $msg.SenderEmailAddress.Length -gt 34) {
            $msg.SenderEmailAddress.Substring(0, 31) + "..."
        } elseif ($msg.SenderEmailAddress) {
            $msg.SenderEmailAddress
        } else {
            "Unknown"
        }

        $date = ConvertTo-SafeDateTime -DateTimeValue $msg.ReceivedDateTime.ToString('yyyy-MM-dd HH:mm')

        Write-Host ($format -f $index, $sender, $email, $date) -ForegroundColor $Global:ColorScheme.Normal
        $index++
    }

    Write-Host ""
    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Function: Invoke-DeleteThread
function Invoke-DeleteThread {
    <#
    .SYNOPSIS
        Deletes entire thread
    .PARAMETER UserEmail
        User email address
    .PARAMETER Thread
        Thread to delete
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [hashtable]$Thread
    )

    Write-Host ""
    $confirm = Show-Confirmation -Message (Get-LocalizedString "thread_confirmDelete" -FormatArgs @($Thread.Messages.Count, $Thread.Subject))

    if (-not $confirm) {
        Write-Host (Get-LocalizedString "performActionAll_deleteCancelled") -ForegroundColor $Global:ColorScheme.Warning
        return
    }

    $successCount = 0
    $errorCount = 0

    foreach ($msg in $Thread.Messages) {
        try {
            Remove-GraphMessage -UserId $UserEmail -MessageId $msg.MessageId | Out-Null
            $successCount++

            Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityDelete") `
                          -Status (Get-LocalizedString "duplicate_progressStatus" -FormatArgs @($successCount, $Thread.Messages.Count)) `
                          -PercentComplete (($successCount / $Thread.Messages.Count) * 100)
        }
        catch {
            Write-Warning "Error deleting message: $($_.Exception.Message)"
            $errorCount++
        }
    }

    Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityDelete") -Completed

    Write-Host ""
    Write-Host (Get-LocalizedString "performActionAll_deleteComplete" -FormatArgs @($successCount)) -ForegroundColor $Global:ColorScheme.Success
    if ($errorCount -gt 0) {
        Write-Host (Get-LocalizedString "performActionAll_deleteErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
    }
}

# Function: Invoke-ArchiveThread
function Invoke-ArchiveThread {
    <#
    .SYNOPSIS
        Archives entire thread
    .PARAMETER UserEmail
        User email address
    .PARAMETER Thread
        Thread to archive
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [hashtable]$Thread
    )

    # Get or create Archive folder
    $folders = Get-GraphMailFolders -UserId $UserEmail
    $archiveFolder = $folders | Where-Object { $_.displayName -eq "Archive" }

    if (-not $archiveFolder) {
        Write-Host (Get-LocalizedString "archiver_creatingFolder") -ForegroundColor $Global:ColorScheme.Info
        $archiveFolder = New-GraphMailFolder -UserId $UserEmail -DisplayName "Archive"
    }

    Write-Host ""
    $confirm = Show-Confirmation -Message (Get-LocalizedString "thread_confirmArchive" -FormatArgs @($Thread.Messages.Count, $Thread.Subject))

    if (-not $confirm) {
        Write-Host (Get-LocalizedString "performActionAll_moveCancelled") -ForegroundColor $Global:ColorScheme.Warning
        return
    }

    $successCount = 0
    $errorCount = 0

    foreach ($msg in $Thread.Messages) {
        try {
            Move-GraphMessage -UserId $UserEmail -MessageId $msg.MessageId -DestinationFolderId $archiveFolder.id | Out-Null
            $successCount++

            Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityMove") `
                          -Status (Get-LocalizedString "duplicate_progressStatus" -FormatArgs @($successCount, $Thread.Messages.Count)) `
                          -PercentComplete (($successCount / $Thread.Messages.Count) * 100)
        }
        catch {
            Write-Warning "Error moving message: $($_.Exception.Message)"
            $errorCount++
        }
    }

    Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityMove") -Completed

    Write-Host ""
    Write-Host (Get-LocalizedString "performActionAll_moveComplete" -FormatArgs @($successCount)) -ForegroundColor $Global:ColorScheme.Success
    if ($errorCount -gt 0) {
        Write-Host (Get-LocalizedString "performActionAll_moveErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
    }
}

# Function: Export-ThreadReport
function Export-ThreadReport {
    <#
    .SYNOPSIS
        Exports thread report to CSV
    .PARAMETER Threads
        Threads hashtable
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Threads
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\thread_report.csv"
    $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    $reportData = @()

    foreach ($threadKey in $Threads.Keys) {
        $thread = $Threads[$threadKey]

        $reportData += [PSCustomObject]@{
            Subject = $thread.Subject
            MessageCount = $thread.Messages.Count
            ParticipantCount = $thread.Participants.Count
            TotalSizeMB = [math]::Round($thread.TotalSize / 1MB, 2)
            OldestMessage = ($thread.Messages | Sort-Object ReceivedDateTime | Select-Object -First 1).ReceivedDateTime
            NewestMessage = ($thread.Messages | Sort-Object ReceivedDateTime -Descending | Select-Object -First 1).ReceivedDateTime
        }
    }

    $reportData | Sort-Object -Property MessageCount -Descending |
        Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
}

# Export functions
Export-ModuleMember -Function Show-ThreadAnalyzer, Get-EmailThreads, Export-ThreadReport
