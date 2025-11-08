<#
.SYNOPSIS
    Email Archiver module for MailCleanBuddy
.DESCRIPTION
    Provides email archiving with retention policies, local backup, and automated cleanup.
    Supports rule-based archiving by age, sender, size, and custom criteria.
#>

# Import dependencies

# Function: Get-EmailsForArchiving
function Get-EmailsForArchiving {
    <#
    .SYNOPSIS
        Gets emails matching archiving criteria
    .PARAMETER AgeMonths
        Archive emails older than X months
    .PARAMETER ExcludeStarred
        Exclude starred/flagged emails
    .PARAMETER ExcludeWithAttachments
        Exclude emails with attachments
    .PARAMETER MinSizeMB
        Only archive emails larger than X MB
    .OUTPUTS
        Array of emails to archive
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$AgeMonths = 6,

        [Parameter(Mandatory = $false)]
        [switch]$ExcludeStarred,

        [Parameter(Mandatory = $false)]
        [switch]$ExcludeWithAttachments,

        [Parameter(Mandatory = $false)]
        [int]$MinSizeMB = 0
    )

    try {
        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            return @()
        }

        $cutoffDate = (Get-Date).AddMonths(-$AgeMonths)
        $minSizeBytes = $MinSizeMB * 1MB
        $emailsToArchive = @()

        foreach ($domain in $cache.Keys) {
            foreach ($msg in $cache[$domain].Messages) {
                # Check age
                if ($msg.ReceivedDateTime) {
                    $receivedDate = ConvertTo-SafeDateTime -DateTimeValue $msg.ReceivedDateTime

                    if ($receivedDate -gt $cutoffDate) {
                        continue  # Too new
                    }
                }

                # Check attachments
                if ($ExcludeWithAttachments -and $msg.HasAttachments) {
                    continue
                }

                # Check size
                if ($MinSizeMB -gt 0 -and $msg.Size -lt $minSizeBytes) {
                    continue
                }

                # Add to archive list
                $emailsToArchive += $msg
            }
        }

        return $emailsToArchive | Sort-Object -Property ReceivedDateTime
    }
    catch {
        Write-Error "Error finding emails for archiving: $($_.Exception.Message)"
        return @()
    }
}

# Function: Show-EmailArchiver
function Show-EmailArchiver {
    <#
    .SYNOPSIS
        Interactive email archiver interface
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
        $title = Get-LocalizedString "archiver_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host (Get-LocalizedString "archiver_description") -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Get archiving criteria
        Write-Host (Get-LocalizedString "archiver_configurePolicy") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ""

        $ageInput = Read-Host (Get-LocalizedString "archiver_enterAge")
        $ageMonths = if ([string]::IsNullOrWhiteSpace($ageInput)) { 6 } else { [int]$ageInput }

        $excludeStarredInput = Read-Host (Get-LocalizedString "archiver_excludeStarred")
        $excludeStarred = $excludeStarredInput -match '^(y|yes|j|ja)$'

        $excludeAttachmentsInput = Read-Host (Get-LocalizedString "archiver_excludeAttachments")
        $excludeWithAttachments = $excludeAttachmentsInput -match '^(y|yes|j|ja)$'

        Write-Host ""
        Write-Host (Get-LocalizedString "archiver_scanning" -FormatArgs @($ageMonths)) -ForegroundColor $Global:ColorScheme.Info

        # Find emails to archive
        $emailsToArchive = Get-EmailsForArchiving -AgeMonths $ageMonths `
                                                   -ExcludeStarred:$excludeStarred `
                                                   -ExcludeWithAttachments:$excludeWithAttachments

        if ($emailsToArchive.Count -eq 0) {
            Write-Host "`n$(Get-LocalizedString 'archiver_noneFound')" -ForegroundColor $Global:ColorScheme.Success
            Write-Host ""
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Calculate statistics
        $totalSize = ($emailsToArchive | Measure-Object -Property Size -Sum).Sum
        $totalSizeMB = [math]::Round($totalSize / 1MB, 2)
        $oldestDate = ($emailsToArchive | Sort-Object ReceivedDateTime | Select-Object -First 1).ReceivedDateTime
        $newestDate = ($emailsToArchive | Sort-Object ReceivedDateTime -Descending | Select-Object -First 1).ReceivedDateTime

        # Display statistics
        Write-Host "`n$(Get-LocalizedString 'archiver_foundCount' -FormatArgs @($emailsToArchive.Count))" -ForegroundColor $Global:ColorScheme.Success
        Write-Host "$(Get-LocalizedString 'archiver_totalSize' -FormatArgs @($totalSizeMB))" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "$(Get-LocalizedString 'archiver_dateRange' -FormatArgs @($oldestDate, $newestDate))" -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Show sample emails
        Write-Host (Get-LocalizedString "archiver_sampleEmails") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        $format = "{0,-45} {1,-25} {2,20}"
        Write-Host ($format -f (Get-LocalizedString 'standardizedList_headerSubject'),
            (Get-LocalizedString 'standardizedList_headerSenderName'),
            (Get-LocalizedString 'standardizedList_headerDate')) -ForegroundColor $Global:ColorScheme.Header
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        $sample = $emailsToArchive | Select-Object -First 10

        foreach ($email in $sample) {
            $subject = if ($email.Subject -and $email.Subject.Length -gt 44) {
                $email.Subject.Substring(0, 41) + "..."
            } elseif ($email.Subject) {
                $email.Subject
            } else {
                Get-LocalizedString 'standardizedList_noSubject'
            }

            $sender = if ($email.SenderName -and $email.SenderName.Length -gt 24) {
                $email.SenderName.Substring(0, 21) + "..."
            } elseif ($email.SenderName) {
                $email.SenderName
            } else {
                "Unknown"
            }

            $date = ConvertTo-SafeDateTime -DateTimeValue $email.ReceivedDateTime.ToString('yyyy-MM-dd HH:mm')

            Write-Host ($format -f $subject, $sender, $date) -ForegroundColor $Global:ColorScheme.Muted
        }

        Write-Host ""

        # Menu for actions
        while ($true) {
            Write-Host (Get-LocalizedString "archiver_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  1. $(Get-LocalizedString 'archiver_exportAndDelete')" -ForegroundColor Yellow
            Write-Host "  2. $(Get-LocalizedString 'archiver_moveToArchive')" -ForegroundColor Cyan
            Write-Host "  3. $(Get-LocalizedString 'archiver_exportOnly')" -ForegroundColor Green
            Write-Host "  4. $(Get-LocalizedString 'archiver_exportReport')" -ForegroundColor Magenta
            Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
            Write-Host ""

            $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

            switch ($choice.ToUpper()) {
                "1" {
                    # Export and delete
                    Invoke-ArchiveAndDelete -UserEmail $UserEmail -Emails $emailsToArchive
                    return
                }
                "2" {
                    # Move to Archive folder
                    Invoke-MoveToArchiveFolder -UserEmail $UserEmail -Emails $emailsToArchive
                    return
                }
                "3" {
                    # Export only
                    Invoke-ExportArchive -UserEmail $UserEmail -Emails $emailsToArchive
                    return
                }
                "4" {
                    # Export report
                    Export-ArchiveReport -Emails $emailsToArchive
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
        Write-Error "Error in email archiver: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Invoke-ArchiveAndDelete
function Invoke-ArchiveAndDelete {
    <#
    .SYNOPSIS
        Exports emails and then deletes them
    .PARAMETER UserEmail
        User email address
    .PARAMETER Emails
        Emails to archive
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Emails
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\email_archive"
    $exportPath = Read-Host (Get-LocalizedString "archiver_enterPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    Write-Host ""
    $confirm = Show-Confirmation -Message (Get-LocalizedString "archiver_confirmExportDelete" -FormatArgs @($Emails.Count))

    if (-not $confirm) {
        Write-Host (Get-LocalizedString "performActionAll_deleteCancelled") -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    # Export to EML format
    Write-Host ""
    Write-Host (Get-LocalizedString "archiver_exporting") -ForegroundColor $Global:ColorScheme.Info

    Export-MessagesAsBatch -UserId $UserEmail -Messages $Emails -ExportPath $exportPath -Format "EML"

    # Delete emails
    Write-Host ""
    Write-Host (Get-LocalizedString "archiver_deleting") -ForegroundColor $Global:ColorScheme.Info

    $successCount = 0
    $errorCount = 0

    foreach ($email in $Emails) {
        try {
            Remove-GraphMessage -UserId $UserEmail -MessageId $email.MessageId | Out-Null
            $successCount++

            Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityDelete") `
                          -Status (Get-LocalizedString "duplicate_progressStatus" -FormatArgs @($successCount, $Emails.Count)) `
                          -PercentComplete (($successCount / $Emails.Count) * 100)
        }
        catch {
            Write-Warning "Error deleting email: $($_.Exception.Message)"
            $errorCount++
        }
    }

    Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityDelete") -Completed

    Write-Host ""
    Write-Host (Get-LocalizedString "archiver_complete" -FormatArgs @($successCount, $exportPath)) -ForegroundColor $Global:ColorScheme.Success
    if ($errorCount -gt 0) {
        Write-Host (Get-LocalizedString "performActionAll_deleteErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
    }

    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Function: Invoke-MoveToArchiveFolder
function Invoke-MoveToArchiveFolder {
    <#
    .SYNOPSIS
        Moves emails to Archive folder in mailbox
    .PARAMETER UserEmail
        User email address
    .PARAMETER Emails
        Emails to move
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Emails
    )

    # Get or create Archive folder
    $folders = Get-GraphMailFolders -UserId $UserEmail
    $archiveFolder = $folders | Where-Object { $_.displayName -eq "Archive" }

    if (-not $archiveFolder) {
        Write-Host (Get-LocalizedString "archiver_creatingFolder") -ForegroundColor $Global:ColorScheme.Info
        $archiveFolder = New-GraphMailFolder -UserId $UserEmail -DisplayName "Archive"
    }

    Write-Host ""
    $confirm = Show-Confirmation -Message (Get-LocalizedString "archiver_confirmMove" -FormatArgs @($Emails.Count))

    if (-not $confirm) {
        Write-Host (Get-LocalizedString "performActionAll_moveCancelled") -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    $successCount = 0
    $errorCount = 0

    foreach ($email in $Emails) {
        try {
            Move-GraphMessage -UserId $UserEmail -MessageId $email.MessageId -DestinationFolderId $archiveFolder.id | Out-Null
            $successCount++

            Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityMove") `
                          -Status (Get-LocalizedString "duplicate_progressStatus" -FormatArgs @($successCount, $Emails.Count)) `
                          -PercentComplete (($successCount / $Emails.Count) * 100)
        }
        catch {
            Write-Warning "Error moving email: $($_.Exception.Message)"
            $errorCount++
        }
    }

    Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityMove") -Completed

    Write-Host ""
    Write-Host (Get-LocalizedString "archiver_moveComplete" -FormatArgs @($successCount)) -ForegroundColor $Global:ColorScheme.Success
    if ($errorCount -gt 0) {
        Write-Host (Get-LocalizedString "performActionAll_moveErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
    }

    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Function: Invoke-ExportArchive
function Invoke-ExportArchive {
    <#
    .SYNOPSIS
        Exports emails only (no delete/move)
    .PARAMETER UserEmail
        User email address
    .PARAMETER Emails
        Emails to export
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Emails
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\email_archive_export"
    $exportPath = Read-Host (Get-LocalizedString "archiver_enterPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "archiver_exporting") -ForegroundColor $Global:ColorScheme.Info

    Export-MessagesAsBatch -UserId $UserEmail -Messages $Emails -ExportPath $exportPath -Format "EML"

    Write-Host ""
    Write-Host (Get-LocalizedString "archiver_exportComplete" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Function: Export-ArchiveReport
function Export-ArchiveReport {
    <#
    .SYNOPSIS
        Exports archive report to CSV
    .PARAMETER Emails
        Emails in archive
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Emails
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\archive_report.csv"
    $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    $reportData = $Emails | Select-Object Subject, SenderName, SenderEmailAddress, ReceivedDateTime,
        @{Name="SizeMB";Expression={[math]::Round($_.Size / 1MB, 2)}}, HasAttachments

    $reportData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Export functions
Export-ModuleMember -Function Show-EmailArchiver, Get-EmailsForArchiving, Export-ArchiveReport
