<#
.SYNOPSIS
    Duplicate Email Detector module for MailCleanBuddy
.DESCRIPTION
    Detects and manages duplicate emails based on multiple criteria including
    subject, sender, date, and content hashing.
#>

# Import dependencies

# Function: Get-EmailHash
function Get-EmailHash {
    <#
    .SYNOPSIS
        Creates a hash for email comparison
    .PARAMETER Message
        Email message object
    .PARAMETER HashType
        Type of hash: 'Quick' (subject+sender+date) or 'Deep' (includes body preview)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Quick', 'Deep')]
        [string]$HashType = 'Quick'
    )

    try {
        # Normalize subject (remove RE:, FW:, etc.)
        $normalizedSubject = if ($Message.Subject) {
            $Message.Subject -replace '^\s*(RE|FW|FWD|AW):\s*', '' -replace '\s+', ' '
        } else {
            ""
        }

        # Get sender email (lowercase)
        $senderEmail = if ($Message.SenderEmailAddress) {
            $Message.SenderEmailAddress.ToLower()
        } else {
            ""
        }

        # Round datetime to nearest minute to catch near-duplicates
        $dateKey = if ($Message.ReceivedDateTime) {
            $date = ConvertTo-SafeDateTime -DateTimeValue $Message.ReceivedDateTime
            $date.ToString('yyyy-MM-dd HH:mm')
        } else {
            ""
        }

        if ($HashType -eq 'Quick') {
            # Quick hash: subject + sender + date (rounded to minute)
            $hashString = "$normalizedSubject|$senderEmail|$dateKey"
        } else {
            # Deep hash: also include body preview if available
            $bodyPreview = if ($Message.BodyPreview) {
                $Message.BodyPreview.Substring(0, [Math]::Min(100, $Message.BodyPreview.Length))
            } else {
                ""
            }
            $hashString = "$normalizedSubject|$senderEmail|$dateKey|$bodyPreview"
        }

        # Create MD5 hash
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $hashBytes = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($hashString))
        $hash = [System.BitConverter]::ToString($hashBytes) -replace '-', ''

        return $hash
    }
    catch {
        Write-Warning "Error creating hash for message: $($_.Exception.Message)"
        return $null
    }
}

# Function: Find-DuplicateEmails
function Find-DuplicateEmails {
    <#
    .SYNOPSIS
        Finds duplicate emails in the cache
    .PARAMETER HashType
        Type of hash to use for comparison
    .OUTPUTS
        Hashtable with duplicates grouped by hash
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('Quick', 'Deep')]
        [string]$HashType = 'Quick'
    )

    try {
        Write-Host (Get-LocalizedString "duplicate_analyzing") -ForegroundColor $Global:ColorScheme.Info

        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            Write-Warning (Get-LocalizedString "analytics_noCacheData")
            return @{}
        }

        # Collect all messages
        $allMessages = @()
        foreach ($domain in $cache.Keys) {
            foreach ($msg in $cache[$domain].Messages) {
                $allMessages += $msg
            }
        }

        Write-Host (Get-LocalizedString "duplicate_processing" -FormatArgs @($allMessages.Count)) -ForegroundColor $Global:ColorScheme.Info

        # Group by hash
        $hashGroups = @{}
        $processedCount = 0

        foreach ($msg in $allMessages) {
            $hash = Get-EmailHash -Message $msg -HashType $HashType

            if ($hash) {
                if (-not $hashGroups.ContainsKey($hash)) {
                    $hashGroups[$hash] = @()
                }
                $hashGroups[$hash] += $msg
            }

            $processedCount++
            if ($processedCount % 100 -eq 0) {
                Write-Progress -Activity (Get-LocalizedString "duplicate_progressActivity") `
                              -Status (Get-LocalizedString "duplicate_progressStatus" -FormatArgs @($processedCount, $allMessages.Count)) `
                              -PercentComplete (($processedCount / $allMessages.Count) * 100)
            }
        }

        Write-Progress -Activity (Get-LocalizedString "duplicate_progressActivity") -Completed

        # Filter to only groups with duplicates (more than 1 message)
        $duplicates = @{}
        foreach ($hash in $hashGroups.Keys) {
            if ($hashGroups[$hash].Count -gt 1) {
                $duplicates[$hash] = $hashGroups[$hash]
            }
        }

        return $duplicates
    }
    catch {
        Write-Error "Error finding duplicates: $($_.Exception.Message)"
        return @{}
    }
}

# Function: Show-DuplicateEmailsManager
function Show-DuplicateEmailsManager {
    <#
    .SYNOPSIS
        Interactive duplicate email management interface
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
        $title = Get-LocalizedString "duplicate_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Find duplicates
        $duplicates = Find-DuplicateEmails -HashType 'Quick'

        if ($duplicates.Count -eq 0) {
            Write-Host "`n$(Get-LocalizedString 'duplicate_noDuplicates')" -ForegroundColor $Global:ColorScheme.Success
            Write-Host ""
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Calculate statistics
        $totalDuplicateGroups = $duplicates.Count
        $totalDuplicateEmails = 0
        $totalSizeBytes = 0

        foreach ($hash in $duplicates.Keys) {
            $group = $duplicates[$hash]
            # Count duplicates (keep one, rest are duplicates)
            $totalDuplicateEmails += ($group.Count - 1)

            # Calculate size
            foreach ($msg in $group) {
                if ($msg.Size) {
                    $totalSizeBytes += $msg.Size
                }
            }
        }

        $potentialSavingsMB = [math]::Round(($totalSizeBytes * ($totalDuplicateEmails / ($totalDuplicateEmails + $totalDuplicateGroups))) / 1MB, 2)

        # Display statistics
        Write-Host "$(Get-LocalizedString 'duplicate_foundGroups' -FormatArgs @($totalDuplicateGroups))" -ForegroundColor $Global:ColorScheme.Success
        Write-Host "$(Get-LocalizedString 'duplicate_foundDuplicates' -FormatArgs @($totalDuplicateEmails))" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "$(Get-LocalizedString 'duplicate_potentialSavings' -FormatArgs @($potentialSavingsMB))" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ""

        # Display top duplicate groups
        Write-Host (Get-LocalizedString "duplicate_topGroups") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        $format = "{0,-4} {1,-50} {2,8} {3,12}"
        Write-Host ($format -f "#", (Get-LocalizedString 'standardizedList_headerSubject'),
            (Get-LocalizedString 'senderOverview_headerCount'), "Size (MB)") -ForegroundColor $Global:ColorScheme.Header
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        # Sort by group size and show top 10
        $sortedGroups = $duplicates.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | Select-Object -First 10
        $index = 1

        foreach ($entry in $sortedGroups) {
            $group = $entry.Value
            $sampleMsg = $group[0]

            $subject = if ($sampleMsg.Subject) {
                if ($sampleMsg.Subject.Length -gt 49) {
                    $sampleMsg.Subject.Substring(0, 46) + "..."
                } else {
                    $sampleMsg.Subject
                }
            } else {
                Get-LocalizedString 'standardizedList_noSubject'
            }

            $groupSize = ($group | Measure-Object -Property Size -Sum).Sum
            $groupSizeMB = [math]::Round($groupSize / 1MB, 2)

            Write-Host ($format -f $index, $subject, $group.Count, $groupSizeMB) -ForegroundColor $Global:ColorScheme.Normal
            $index++
        }

        Write-Host ""

        # Menu for actions
        while ($true) {
            Write-Host (Get-LocalizedString "duplicate_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  1. $(Get-LocalizedString 'duplicate_viewGroup')" -ForegroundColor Green
            Write-Host "  2. $(Get-LocalizedString 'duplicate_autoClean')" -ForegroundColor Yellow
            Write-Host "  3. $(Get-LocalizedString 'duplicate_exportReport')" -ForegroundColor Cyan
            Write-Host "  4. $(Get-LocalizedString 'duplicate_deepScan')" -ForegroundColor Magenta
            Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
            Write-Host ""

            $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

            switch ($choice.ToUpper()) {
                "1" {
                    # View specific duplicate group
                    $groupNum = Read-Host (Get-LocalizedString "duplicate_enterGroupNumber")
                    if ($groupNum -match '^\d+$' -and [int]$groupNum -ge 1 -and [int]$groupNum -le $sortedGroups.Count) {
                        $selectedGroup = $sortedGroups[[int]$groupNum - 1].Value
                        Show-DuplicateGroupDetails -UserEmail $UserEmail -Group $selectedGroup
                    } else {
                        Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
                    }
                }
                "2" {
                    # Auto-clean duplicates
                    Invoke-AutoCleanDuplicates -UserEmail $UserEmail -Duplicates $duplicates
                    # Refresh after cleaning
                    return
                }
                "3" {
                    # Export report
                    Export-DuplicateReport -Duplicates $duplicates
                }
                "4" {
                    # Deep scan (includes body content)
                    Write-Host ""
                    Write-Host (Get-LocalizedString "duplicate_deepScanWarning") -ForegroundColor $Global:ColorScheme.Warning
                    $confirm = Show-Confirmation -Message (Get-LocalizedString "duplicate_deepScanConfirm")
                    if ($confirm) {
                        $duplicates = Find-DuplicateEmails -HashType 'Deep'
                        return
                    }
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
        Write-Error "Error in duplicate manager: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Show-DuplicateGroupDetails
function Show-DuplicateGroupDetails {
    <#
    .SYNOPSIS
        Shows details of a duplicate group
    .PARAMETER UserEmail
        User email address
    .PARAMETER Group
        Array of duplicate messages
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Group
    )

    Clear-Host
    Write-Host "`n$(Get-LocalizedString 'duplicate_groupDetails' -FormatArgs @($Group.Count))" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    # Show sample message details
    $sample = $Group[0]
    Write-Host "  $(Get-LocalizedString 'standardizedList_headerSubject'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $sample.Subject -ForegroundColor $Global:ColorScheme.Value

    Write-Host "  $(Get-LocalizedString 'standardizedList_headerSenderName'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host "$($sample.SenderName) <$($sample.SenderEmailAddress)>" -ForegroundColor $Global:ColorScheme.Value

    Write-Host ""
    Write-Host (Get-LocalizedString "duplicate_instances") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

    $format = "{0,-4} {1,-25} {2,12} {3,10}"
    Write-Host ($format -f "#", (Get-LocalizedString 'standardizedList_headerDate'),
        "Size (KB)", "Status") -ForegroundColor $Global:ColorScheme.Header
    Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

    # Sort by date (newest first)
    $sortedGroup = $Group | Sort-Object { ConvertTo-SafeDateTime -DateTimeValue $_.ReceivedDateTime } -Descending
    $index = 1

    foreach ($msg in $sortedGroup) {
        $date = ConvertTo-SafeDateTime -DateTimeValue $msg.ReceivedDateTime.ToString('yyyy-MM-dd HH:mm')
        $sizeKB = [math]::Round($msg.Size / 1KB, 1)

        $status = if ($index -eq 1) {
            Get-LocalizedString 'duplicate_keepNewest'
        } else {
            Get-LocalizedString 'duplicate_canDelete'
        }

        $color = if ($index -eq 1) { $Global:ColorScheme.Success } else { $Global:ColorScheme.Muted }
        Write-Host ($format -f $index, $date, $sizeKB, $status) -ForegroundColor $color
        $index++
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "duplicate_autoKeepNewest") -ForegroundColor $Global:ColorScheme.Info
    Write-Host ""

    # Ask what to do
    Write-Host "1. $(Get-LocalizedString 'duplicate_deleteOlder')" -ForegroundColor Yellow
    Write-Host "2. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
    Write-Host ""

    $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

    if ($choice -eq "1") {
        # Delete all except the newest
        $toDelete = $sortedGroup | Select-Object -Skip 1

        $confirm = Show-Confirmation -Message (Get-LocalizedString "duplicate_confirmDelete" -FormatArgs @($toDelete.Count))

        if ($confirm) {
            $successCount = 0
            $errorCount = 0

            foreach ($msg in $toDelete) {
                try {
                    Remove-GraphMessage -UserId $UserEmail -MessageId $msg.MessageId | Out-Null
                    $successCount++

                    Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityDelete") `
                                  -Status (Get-LocalizedString "performActionAll_progressStatusDelete" -FormatArgs @($msg.Subject)) `
                                  -PercentComplete (($successCount / $toDelete.Count) * 100)
                }
                catch {
                    Write-Warning (Get-LocalizedString "performActionAll_errorDeletingEmailId" -FormatArgs @($msg.MessageId, $_.Exception.Message))
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
    }
}

# Function: Invoke-AutoCleanDuplicates
function Invoke-AutoCleanDuplicates {
    <#
    .SYNOPSIS
        Automatically cleans duplicates by keeping the newest copy
    .PARAMETER UserEmail
        User email address
    .PARAMETER Duplicates
        Hashtable of duplicate groups
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [hashtable]$Duplicates
    )

    try {
        # Calculate total to delete
        $totalToDelete = 0
        foreach ($hash in $Duplicates.Keys) {
            $totalToDelete += ($Duplicates[$hash].Count - 1)
        }

        Write-Host ""
        Write-Host (Get-LocalizedString "duplicate_autoCleanInfo" -FormatArgs @($totalToDelete)) -ForegroundColor $Global:ColorScheme.Info
        Write-Host (Get-LocalizedString "duplicate_autoCleanStrategy") -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        $confirm = Show-Confirmation -Message (Get-LocalizedString "duplicate_autoCleanConfirm" -FormatArgs @($totalToDelete))

        if (-not $confirm) {
            Write-Host (Get-LocalizedString "performActionAll_deleteCancelled") -ForegroundColor $Global:ColorScheme.Warning
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        Write-Host ""
        Write-Host (Get-LocalizedString "duplicate_autoCleanStarting") -ForegroundColor $Global:ColorScheme.Info

        $successCount = 0
        $errorCount = 0
        $processedGroups = 0

        foreach ($hash in $Duplicates.Keys) {
            $group = $Duplicates[$hash]
            $processedGroups++

            # Sort by date (newest first) and skip the first one
            $sortedGroup = $group | Sort-Object { ConvertTo-SafeDateTime -DateTimeValue $_.ReceivedDateTime } -Descending
            $toDelete = $sortedGroup | Select-Object -Skip 1

            foreach ($msg in $toDelete) {
                try {
                    Remove-GraphMessage -UserId $UserEmail -MessageId $msg.MessageId | Out-Null
                    $successCount++

                    Write-Progress -Activity (Get-LocalizedString "duplicate_autoCleanProgress") `
                                  -Status (Get-LocalizedString "duplicate_progressStatus" -FormatArgs @($successCount, $totalToDelete)) `
                                  -PercentComplete (($successCount / $totalToDelete) * 100)
                }
                catch {
                    Write-Warning (Get-LocalizedString "performActionAll_errorDeletingEmailId" -FormatArgs @($msg.MessageId, $_.Exception.Message))
                    $errorCount++
                }
            }
        }

        Write-Progress -Activity (Get-LocalizedString "duplicate_autoCleanProgress") -Completed

        Write-Host ""
        Write-Host (Get-LocalizedString "duplicate_autoCleanComplete" -FormatArgs @($successCount, $processedGroups)) -ForegroundColor $Global:ColorScheme.Success
        if ($errorCount -gt 0) {
            Write-Host (Get-LocalizedString "performActionAll_deleteErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
        }

        Write-Host ""
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
    catch {
        Write-Error "Error in auto-clean: $($_.Exception.Message)"
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Export-DuplicateReport
function Export-DuplicateReport {
    <#
    .SYNOPSIS
        Exports duplicate report to CSV
    .PARAMETER Duplicates
        Hashtable of duplicate groups
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Duplicates
    )

    try {
        $defaultPath = Join-Path $PSScriptRoot "..\..\duplicate_report.csv"
        $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

        if ([string]::IsNullOrWhiteSpace($exportPath)) {
            $exportPath = $defaultPath
        }

        # Create report data
        $reportData = @()

        foreach ($hash in $Duplicates.Keys) {
            $group = $Duplicates[$hash]
            $sample = $group[0]

            $groupSize = ($group | Measure-Object -Property Size -Sum).Sum
            $groupSizeMB = [math]::Round($groupSize / 1MB, 2)

            $reportData += [PSCustomObject]@{
                Subject = $sample.Subject
                Sender = $sample.SenderEmailAddress
                Count = $group.Count
                TotalSizeMB = $groupSizeMB
                OldestDate = ($group | Sort-Object ReceivedDateTime | Select-Object -First 1).ReceivedDateTime
                NewestDate = ($group | Sort-Object ReceivedDateTime -Descending | Select-Object -First 1).ReceivedDateTime
            }
        }

        # Export to CSV
        $reportData | Sort-Object -Property Count -Descending |
            Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

        Write-Host ""
        Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
    catch {
        Write-Error "Error exporting report: $($_.Exception.Message)"
        Write-Host (Get-LocalizedString "unsubscribe_exportError" -FormatArgs @($_.Exception.Message)) -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Export functions
Export-ModuleMember -Function Show-DuplicateEmailsManager, Find-DuplicateEmails, `
    Get-EmailHash, Export-DuplicateReport
