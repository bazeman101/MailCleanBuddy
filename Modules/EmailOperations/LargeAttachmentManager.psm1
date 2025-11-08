<#
.SYNOPSIS
    Large Attachment Manager module for MailCleanBuddy
.DESCRIPTION
    Identifies and manages emails with large attachments to optimize mailbox storage.
    Provides bulk download, delete, and storage analysis features.
#>

# Import dependencies

# Function: Get-LargeAttachments
function Get-LargeAttachments {
    <#
    .SYNOPSIS
        Finds emails with large attachments
    .PARAMETER MinSizeMB
        Minimum attachment size in MB (default: 5)
    .OUTPUTS
        Array of emails with large attachments
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$MinSizeMB = 5
    )

    try {
        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            return @()
        }

        $minSizeBytes = $MinSizeMB * 1MB
        $largeAttachments = @()

        foreach ($domain in $cache.Keys) {
            foreach ($msg in $cache[$domain].Messages) {
                # Check if has attachments and meets size requirement
                if ($msg.HasAttachments -and $msg.Size -ge $minSizeBytes) {
                    $largeAttachments += $msg
                }
            }
        }

        return $largeAttachments | Sort-Object -Property Size -Descending
    }
    catch {
        Write-Error "Error finding large attachments: $($_.Exception.Message)"
        return @()
    }
}

# Function: Show-LargeAttachmentManager
function Show-LargeAttachmentManager {
    <#
    .SYNOPSIS
        Interactive large attachment management interface
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
        $title = Get-LocalizedString "largeAttach_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Ask for minimum size
        Write-Host (Get-LocalizedString "largeAttach_enterMinSize") -ForegroundColor $Global:ColorScheme.Info
        $minSizeInput = Read-Host "Minimum size in MB (default: 5)"
        $minSize = if ([string]::IsNullOrWhiteSpace($minSizeInput)) { 5 } else { [int]$minSizeInput }

        Write-Host ""
        Write-Host (Get-LocalizedString "largeAttach_scanning" -FormatArgs @($minSize)) -ForegroundColor $Global:ColorScheme.Info

        # Find large attachments
        $largeAttachments = Get-LargeAttachments -MinSizeMB $minSize

        if ($largeAttachments.Count -eq 0) {
            Write-Host "`n$(Get-LocalizedString 'largeAttach_noneFound' -FormatArgs @($minSize))" -ForegroundColor $Global:ColorScheme.Success
            Write-Host ""
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Calculate statistics
        $totalSize = ($largeAttachments | Measure-Object -Property Size -Sum).Sum
        $totalSizeMB = [math]::Round($totalSize / 1MB, 2)
        $avgSizeMB = [math]::Round($totalSizeMB / $largeAttachments.Count, 2)

        # Display statistics
        Write-Host "`n$(Get-LocalizedString 'largeAttach_foundCount' -FormatArgs @($largeAttachments.Count))" -ForegroundColor $Global:ColorScheme.Success
        Write-Host "$(Get-LocalizedString 'largeAttach_totalSize' -FormatArgs @($totalSizeMB))" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "$(Get-LocalizedString 'largeAttach_avgSize' -FormatArgs @($avgSizeMB))" -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Display top emails
        Write-Host (Get-LocalizedString "largeAttach_topEmails") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        $format = "{0,-4} {1,-40} {2,-25} {3,12}"
        Write-Host ($format -f "#", (Get-LocalizedString 'standardizedList_headerSubject'),
            (Get-LocalizedString 'standardizedList_headerSenderName'), "Size (MB)") -ForegroundColor $Global:ColorScheme.Header
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        # Show top 15
        $topAttachments = $largeAttachments | Select-Object -First 15
        $index = 1

        foreach ($email in $topAttachments) {
            $subject = if ($email.Subject) {
                if ($email.Subject.Length -gt 39) {
                    $email.Subject.Substring(0, 36) + "..."
                } else {
                    $email.Subject
                }
            } else {
                Get-LocalizedString 'standardizedList_noSubject'
            }

            $sender = if ($email.SenderName) {
                if ($email.SenderName.Length -gt 24) {
                    $email.SenderName.Substring(0, 21) + "..."
                } else {
                    $email.SenderName
                }
            } else {
                "Unknown"
            }

            $sizeMB = [math]::Round($email.Size / 1MB, 2)

            Write-Host ($format -f $index, $subject, $sender, $sizeMB) -ForegroundColor $Global:ColorScheme.Normal
            $index++
        }

        Write-Host ""

        # Menu for actions
        while ($true) {
            Write-Host (Get-LocalizedString "largeAttach_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  1. $(Get-LocalizedString 'largeAttach_downloadAndDelete')" -ForegroundColor Yellow
            Write-Host "  2. $(Get-LocalizedString 'largeAttach_downloadOnly')" -ForegroundColor Green
            Write-Host "  3. $(Get-LocalizedString 'largeAttach_deleteOnly')" -ForegroundColor Red
            Write-Host "  4. $(Get-LocalizedString 'largeAttach_exportReport')" -ForegroundColor Cyan
            Write-Host "  5. $(Get-LocalizedString 'largeAttach_filterByType')" -ForegroundColor Magenta
            Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
            Write-Host ""

            $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

            switch ($choice.ToUpper()) {
                "1" {
                    # Download and delete
                    Invoke-DownloadAndDeleteLargeAttachments -UserEmail $UserEmail -Emails $largeAttachments
                    return
                }
                "2" {
                    # Download only
                    Invoke-DownloadLargeAttachments -UserEmail $UserEmail -Emails $largeAttachments
                }
                "3" {
                    # Delete only
                    Invoke-DeleteEmailsWithLargeAttachments -UserEmail $UserEmail -Emails $largeAttachments
                    return
                }
                "4" {
                    # Export report
                    Export-LargeAttachmentReport -Emails $largeAttachments
                }
                "5" {
                    # Filter by type
                    Show-FilteredByType -UserEmail $UserEmail -Emails $largeAttachments
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
        Write-Error "Error in large attachment manager: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Invoke-DownloadAndDeleteLargeAttachments
function Invoke-DownloadAndDeleteLargeAttachments {
    <#
    .SYNOPSIS
        Downloads attachments and then deletes the emails
    .PARAMETER UserEmail
        User email address
    .PARAMETER Emails
        Array of emails to process
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Emails
    )

    try {
        $defaultPath = Join-Path $PSScriptRoot "..\..\large_attachments_backup"
        $savePath = Read-Host (Get-LocalizedString "largeAttach_enterPath" -FormatArgs @($defaultPath))

        if ([string]::IsNullOrWhiteSpace($savePath)) {
            $savePath = $defaultPath
        }

        # Ensure path exists
        if (-not (Test-Path $savePath)) {
            New-Item -Path $savePath -ItemType Directory -Force | Out-Null
        }

        Write-Host ""
        $confirm = Show-Confirmation -Message (Get-LocalizedString "largeAttach_confirmDownloadDelete" -FormatArgs @($Emails.Count))

        if (-not $confirm) {
            Write-Host (Get-LocalizedString "performActionAll_deleteCancelled") -ForegroundColor $Global:ColorScheme.Warning
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        Write-Host ""
        Write-Host (Get-LocalizedString "largeAttach_startingDownload") -ForegroundColor $Global:ColorScheme.Info

        $successCount = 0
        $errorCount = 0

        foreach ($email in $Emails) {
            try {
                # Create folder for this email
                $emailFolder = Join-Path $savePath ($email.Subject -replace '[\\/:*?"<>|]', '_')
                if (-not (Test-Path $emailFolder)) {
                    New-Item -Path $emailFolder -ItemType Directory -Force | Out-Null
                }

                # Download attachments
                $attachments = Get-MgUserMessageAttachment -UserId $UserEmail -MessageId $email.MessageId

                foreach ($attachment in $attachments) {
                    $attachmentPath = Join-Path $emailFolder $attachment.Name
                    $content = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserEmail/messages/$($email.MessageId)/attachments/$($attachment.Id)/`$value"

                    if ($content) {
                        [System.IO.File]::WriteAllBytes($attachmentPath, $content)
                    }
                }

                # Delete email after successful download
                Remove-GraphMessage -UserId $UserEmail -MessageId $email.MessageId | Out-Null
                $successCount++

                Write-Progress -Activity (Get-LocalizedString "largeAttach_downloadProgress") `
                              -Status (Get-LocalizedString "duplicate_progressStatus" -FormatArgs @($successCount, $Emails.Count)) `
                              -PercentComplete (($successCount / $Emails.Count) * 100)
            }
            catch {
                Write-Warning "Error processing email '$($email.Subject)': $($_.Exception.Message)"
                $errorCount++
            }
        }

        Write-Progress -Activity (Get-LocalizedString "largeAttach_downloadProgress") -Completed

        Write-Host ""
        Write-Host (Get-LocalizedString "largeAttach_downloadComplete" -FormatArgs @($successCount, $savePath)) -ForegroundColor $Global:ColorScheme.Success
        if ($errorCount -gt 0) {
            Write-Host (Get-LocalizedString "performActionAll_deleteErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
        }

        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
    catch {
        Write-Error "Error in download and delete: $($_.Exception.Message)"
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Invoke-DownloadLargeAttachments
function Invoke-DownloadLargeAttachments {
    <#
    .SYNOPSIS
        Downloads attachments only (keeps emails)
    .PARAMETER UserEmail
        User email address
    .PARAMETER Emails
        Array of emails to process
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Emails
    )

    try {
        $defaultPath = Join-Path $PSScriptRoot "..\..\large_attachments"
        $savePath = Read-Host (Get-LocalizedString "largeAttach_enterPath" -FormatArgs @($defaultPath))

        if ([string]::IsNullOrWhiteSpace($savePath)) {
            $savePath = $defaultPath
        }

        # Ensure path exists
        if (-not (Test-Path $savePath)) {
            New-Item -Path $savePath -ItemType Directory -Force | Out-Null
        }

        Write-Host ""
        Write-Host (Get-LocalizedString "largeAttach_startingDownload") -ForegroundColor $Global:ColorScheme.Info

        $successCount = 0
        $errorCount = 0
        $totalDownloaded = 0

        foreach ($email in $Emails) {
            try {
                # Get message ID
                $msgId = if ($email.MessageId) { $email.MessageId } elseif ($email.Id) { $email.Id } else { $null }
                if (-not $msgId) {
                    Write-Warning "Skipping email without message ID"
                    continue
                }

                # Download attachments for this email
                Write-Host "Processing: $($email.Subject)..." -ForegroundColor Cyan
                $attachments = Get-MgUserMessageAttachment -UserId $UserEmail -MessageId $msgId -ErrorAction Stop

                if ($attachments -and $attachments.Count -gt 0) {
                    foreach ($attachment in $attachments) {
                        try {
                            # Create safe filename
                            $safeFilename = $attachment.Name -replace '[\\/:*?"<>|]', '_'
                            $attachmentPath = Join-Path $savePath $safeFilename

                            # Make filename unique if it exists
                            if (Test-Path $attachmentPath) {
                                $extension = [System.IO.Path]::GetExtension($safeFilename)
                                $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($safeFilename)
                                $counter = 1
                                do {
                                    $attachmentPath = Join-Path $savePath "$nameWithoutExt`_$counter$extension"
                                    $counter++
                                } while (Test-Path $attachmentPath)
                            }

                            # Get attachment content
                            $content = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserEmail/messages/$msgId/attachments/$($attachment.Id)/`$value" -ErrorAction Stop

                            if ($content) {
                                [System.IO.File]::WriteAllBytes($attachmentPath, $content)
                                Write-Host "  Downloaded: $($attachment.Name)" -ForegroundColor Green
                                $totalDownloaded++
                            }
                        }
                        catch {
                            Write-Warning "  Failed to download $($attachment.Name): $($_.Exception.Message)"
                            $errorCount++
                        }
                    }
                }

                $successCount++
                Write-Progress -Activity (Get-LocalizedString "largeAttach_downloadProgress") `
                              -Status (Get-LocalizedString "duplicate_progressStatus" -FormatArgs @($successCount, $Emails.Count)) `
                              -PercentComplete (($successCount / $Emails.Count) * 100)
            }
            catch {
                Write-Warning "Error processing email '$($email.Subject)': $($_.Exception.Message)"
                $errorCount++
            }
        }

        Write-Progress -Activity (Get-LocalizedString "largeAttach_downloadProgress") -Completed

        Write-Host ""
        Write-Host "Downloaded $totalDownloaded attachment(s) to: $savePath" -ForegroundColor $Global:ColorScheme.Success
        if ($errorCount -gt 0) {
            Write-Host "Errors: $errorCount" -ForegroundColor $Global:ColorScheme.Warning
        }

        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
    catch {
        Write-Error "Error in download: $($_.Exception.Message)"
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Invoke-DeleteEmailsWithLargeAttachments
function Invoke-DeleteEmailsWithLargeAttachments {
    <#
    .SYNOPSIS
        Deletes emails with large attachments (no download)
    .PARAMETER UserEmail
        User email address
    .PARAMETER Emails
        Array of emails to delete
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Emails
    )

    Write-Host ""
    $confirm = Show-Confirmation -Message (Get-LocalizedString "largeAttach_confirmDelete" -FormatArgs @($Emails.Count))

    if (-not $confirm) {
        Write-Host (Get-LocalizedString "performActionAll_deleteCancelled") -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

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
    Write-Host (Get-LocalizedString "performActionAll_deleteComplete" -FormatArgs @($successCount)) -ForegroundColor $Global:ColorScheme.Success
    if ($errorCount -gt 0) {
        Write-Host (Get-LocalizedString "performActionAll_deleteErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
    }

    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Function: Export-LargeAttachmentReport
function Export-LargeAttachmentReport {
    <#
    .SYNOPSIS
        Exports report to CSV
    .PARAMETER Emails
        Array of emails with large attachments
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Emails
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\large_attachments_report.csv"
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

# Function: Show-FilteredByType
function Show-FilteredByType {
    <#
    .SYNOPSIS
        Shows attachments filtered by file type
    .PARAMETER UserEmail
        User email address
    .PARAMETER Emails
        Array of emails
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Emails
    )

    Write-Host ""
    Write-Host (Get-LocalizedString "largeAttach_enterFileTypes") -ForegroundColor $Global:ColorScheme.Info
    $fileTypesInput = Read-Host "File types (e.g., 'pdf,zip,mp4')"

    if ([string]::IsNullOrWhiteSpace($fileTypesInput)) {
        Write-Host (Get-LocalizedString "unsubscribe_invalidChoice") -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    $fileTypes = $fileTypesInput -split ',' | ForEach-Object { $_.Trim().ToLower() }
    Write-Host (Get-LocalizedString "largeAttach_filteringByType" -FormatArgs @($fileTypesInput)) -ForegroundColor $Global:ColorScheme.Info

    # This would require fetching actual attachment details - simplified for now
    Write-Host (Get-LocalizedString "largeAttach_featureComingSoon") -ForegroundColor $Global:ColorScheme.Warning
    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Export functions
Export-ModuleMember -Function Show-LargeAttachmentManager, Get-LargeAttachments, `
    Export-LargeAttachmentReport
