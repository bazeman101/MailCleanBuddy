<#
.SYNOPSIS
    Email List View module for MailCleanBuddy
.DESCRIPTION
    Provides email list viewing and basic email action functionality
#>

# Import dependencies

# Function: Show-StandardizedEmailListView
function Show-StandardizedEmailListView {
    <#
    .SYNOPSIS
        Displays a standardized email list view with navigation and actions
    .PARAMETER UserEmail
        User email address
    .PARAMETER Messages
        Array of messages to display
    .PARAMETER Title
        View title
    .PARAMETER AllowActions
        Whether to allow actions (default: true)
    .PARAMETER ViewName
        View identifier name
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Messages,

        [Parameter(Mandatory = $false)]
        [string]$Title = "Email List",

        [Parameter(Mandatory = $false)]
        [bool]$AllowActions = $true,

        [Parameter(Mandatory = $false)]
        [string]$ViewName = "EmailList"
    )

    if (-not $Messages -or $Messages.Count -eq 0) {
        Clear-Host
        Write-Host "`n$Title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""
        Write-Host (Get-LocalizedString "standardizedList_noEmailsToDisplay") -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    $currentMessages = $Messages
    $selectedEmailIndex = 0
    $topDisplayIndex = 0
    $displayLines = [Math]::Max(10, $Host.UI.RawUI.WindowSize.Height - 10)
    $spaceSelectedMessageIds = [System.Collections.Generic.HashSet[string]]::new()

    $emailListLoopActive = $true
    while ($emailListLoopActive) {
        Clear-Host

        # Display header
        Write-Host "`n$Title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        if ($AllowActions) {
            Write-Host "Actions: [Enter] View/Actions | [H] Header Analysis | [Space] Select | [A] Select All | [N] Deselect All" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "         [Del] Delete | [V] Move | [Q/Esc] Back" -ForegroundColor $Global:ColorScheme.Info
        } else {
            Write-Host "Navigation: [Up/Down] Navigate | [Q/Esc] Back" -ForegroundColor $Global:ColorScheme.Info
        }
        Write-Host ""

        # Column headers
        $headerFormat = "{0,-5} {1,-20} {2,-25} {3,-40}"
        Write-Host ($headerFormat -f "#", "Date", "From", "Subject") -ForegroundColor $Global:ColorScheme.Header
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

        # Display messages
        $currentDisplayLines = [Math]::Min($displayLines, $currentMessages.Count)
        $endDisplayIndex = [Math]::Min(($topDisplayIndex + $currentDisplayLines - 1), ($currentMessages.Count - 1))

        for ($i = $topDisplayIndex; $i -le $endDisplayIndex; $i++) {
            if ($i -ge $currentMessages.Count) { break }
            $message = $currentMessages[$i]
            $itemNumber = $i + 1

            $receivedDisplay = Format-SafeDateTime -DateTimeValue $message.ReceivedDateTime -ShortFormat

            $senderDisplay = if ($message.SenderEmailAddress) {
                $message.SenderEmailAddress
            } elseif ($message.SenderName) {
                $message.SenderName
            } else { "N/A" }
            if ($senderDisplay.Length -gt 24) { $senderDisplay = $senderDisplay.Substring(0, 21) + "..." }

            $subjectDisplay = if ($message.Subject) { $message.Subject } else { "(No Subject)" }
            if ($subjectDisplay.Length -gt 39) { $subjectDisplay = $subjectDisplay.Substring(0, 36) + "..." }

            $selectionPrefix = "   "
            $currentLineFgColor = $Global:ColorScheme.Normal

            if ($spaceSelectedMessageIds.Contains($message.Id)) {
                $selectionPrefix = "[*]"
            }

            if ($i -eq $selectedEmailIndex) {
                $selectionPrefix = if ($selectionPrefix -match "\[\*\]") { ">*]" } else { ">  " }
                $currentLineFgColor = $Global:ColorScheme.Highlight
            }

            $lineText = "{0} {1,-5} {2,-20} {3,-24} {4,-39}" -f $selectionPrefix, "$itemNumber.", $receivedDisplay, $senderDisplay, $subjectDisplay

            Write-Host $lineText -ForegroundColor $currentLineFgColor
        }

        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        $shownCountStart = if($currentMessages.Count -gt 0) { $topDisplayIndex + 1 } else { 0 }
        $shownCountEnd = if($currentMessages.Count -gt 0) { $endDisplayIndex + 1 } else { 0 }
        Write-Host "Showing $shownCountStart-$shownCountEnd of $($currentMessages.Count) | Selected: $($spaceSelectedMessageIds.Count)" -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Read key
        $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

        switch ($keyInfo.VirtualKeyCode) {
            38 { # UpArrow
                if ($selectedEmailIndex -gt 0) { $selectedEmailIndex-- }
                if ($selectedEmailIndex -lt $topDisplayIndex) { $topDisplayIndex = $selectedEmailIndex }
            }
            40 { # DownArrow
                if ($currentMessages.Count -gt 0 -and $selectedEmailIndex -lt ($currentMessages.Count - 1)) { $selectedEmailIndex++ }
                if ($selectedEmailIndex -gt $endDisplayIndex -and $topDisplayIndex -lt ($currentMessages.Count - $currentDisplayLines)) { $topDisplayIndex++ }
            }
            33 { # PageUp
                $selectedEmailIndex = [Math]::Max(0, $selectedEmailIndex - $currentDisplayLines)
                $topDisplayIndex = [Math]::Max(0, $topDisplayIndex - $currentDisplayLines)
                if ($selectedEmailIndex -lt $topDisplayIndex) {$topDisplayIndex = $selectedEmailIndex}
            }
            34 { # PageDown
                if ($currentMessages.Count -gt 0) {
                    $selectedEmailIndex = [Math]::Min(($currentMessages.Count - 1), $selectedEmailIndex + $currentDisplayLines)
                    $topDisplayIndex = [Math]::Min(($currentMessages.Count - $currentDisplayLines), $topDisplayIndex + $currentDisplayLines)
                    if ($topDisplayIndex -lt 0) {$topDisplayIndex = 0}
                }
            }
            32 { # Spacebar
                if ($AllowActions -and $currentMessages.Count -gt 0 -and $selectedEmailIndex -ge 0 -and $selectedEmailIndex -lt $currentMessages.Count) {
                    $currentMessageIdForSpace = $currentMessages[$selectedEmailIndex].Id
                    if ($spaceSelectedMessageIds.Contains($currentMessageIdForSpace)) {
                        $spaceSelectedMessageIds.Remove($currentMessageIdForSpace) | Out-Null
                    } else {
                        $spaceSelectedMessageIds.Add($currentMessageIdForSpace) | Out-Null
                    }
                }
            }
            13 { # Enter - View details and actions
                if ($currentMessages.Count -gt 0 -and $selectedEmailIndex -ge 0 -and $selectedEmailIndex -lt $currentMessages.Count) {
                    $selectedMessage = $currentMessages[$selectedEmailIndex]

                    # Determine the message ID - check both Id and MessageId properties
                    # Handle both hashtables (from cache) and PSCustomObjects
                    $messageId = $null
                    if (-not [string]::IsNullOrWhiteSpace($selectedMessage.Id)) {
                        $messageId = $selectedMessage.Id
                    } elseif (-not [string]::IsNullOrWhiteSpace($selectedMessage.MessageId)) {
                        $messageId = $selectedMessage.MessageId
                    }

                    if (-not [string]::IsNullOrWhiteSpace($messageId)) {
                        Show-EmailActionsMenu -UserEmail $UserEmail -MessageId $messageId -AllMessages $currentMessages -CurrentIndex $selectedEmailIndex
                        # Refresh the message list to reflect any changes (deleted/moved emails)
                        # Note: This is a simplified refresh - in production you'd reload from source
                    } else {
                        Write-Host "`nError: Could not find message ID" -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
            }
            27 { $emailListLoopActive = $false } # Escape
            46 { # Delete
                if ($AllowActions) {
                    $messagesToDelete = @()
                    if ($spaceSelectedMessageIds.Count -gt 0) {
                        $messagesToDelete = $currentMessages | Where-Object { $spaceSelectedMessageIds.Contains($_.Id) }
                    } elseif ($currentMessages.Count -gt 0 -and $selectedEmailIndex -ge 0 -and $selectedEmailIndex -lt $currentMessages.Count) {
                        $messagesToDelete = @($currentMessages[$selectedEmailIndex])
                    }

                    if ($messagesToDelete.Count -gt 0) {
                        $confirm = Show-Confirmation -Message "Delete $($messagesToDelete.Count) email(s)?"
                        if ($confirm) {
                            foreach ($msg in $messagesToDelete) {
                                try {
                                    Remove-GraphMessage -UserId $UserEmail -MessageId $msg.Id | Out-Null
                                } catch {
                                    Write-Warning "Failed to delete message: $($_.Exception.Message)"
                                }
                            }
                            Write-Host "$($messagesToDelete.Count) email(s) deleted." -ForegroundColor $Global:ColorScheme.Success
                            Start-Sleep -Seconds 1

                            # Remove deleted messages from list
                            $deletedIds = $messagesToDelete | ForEach-Object { $_.Id }
                            $currentMessages = $currentMessages | Where-Object { $deletedIds -notcontains $_.Id }
                            $spaceSelectedMessageIds.Clear()

                            if ($currentMessages.Count -eq 0) {
                                $emailListLoopActive = $false
                            } else {
                                $selectedEmailIndex = [Math]::Min($selectedEmailIndex, $currentMessages.Count - 1)
                            }
                        }
                    }
                }
            }
            86 { # V - Move
                if ($AllowActions) {
                    $messagesToMove = @()
                    if ($spaceSelectedMessageIds.Count -gt 0) {
                        $messagesToMove = $currentMessages | Where-Object { $spaceSelectedMessageIds.Contains($_.Id) }
                    } elseif ($currentMessages.Count -gt 0 -and $selectedEmailIndex -ge 0 -and $selectedEmailIndex -lt $currentMessages.Count) {
                        $messagesToMove = @($currentMessages[$selectedEmailIndex])
                    }

                    if ($messagesToMove.Count -gt 0) {
                        $folder = Select-MailFolder -UserEmail $UserEmail
                        if ($folder) {
                            foreach ($msg in $messagesToMove) {
                                try {
                                    Move-GraphMessage -UserId $UserEmail -MessageId $msg.Id -DestinationFolderId $folder | Out-Null
                                } catch {
                                    Write-Warning "Failed to move message: $($_.Exception.Message)"
                                }
                            }
                            Write-Host "$($messagesToMove.Count) email(s) moved." -ForegroundColor $Global:ColorScheme.Success
                            Start-Sleep -Seconds 1

                            # Remove moved messages from list
                            $movedIds = $messagesToMove | ForEach-Object { $_.Id }
                            $currentMessages = $currentMessages | Where-Object { $movedIds -notcontains $_.Id }
                            $spaceSelectedMessageIds.Clear()

                            if ($currentMessages.Count -eq 0) {
                                $emailListLoopActive = $false
                            } else {
                                $selectedEmailIndex = [Math]::Min($selectedEmailIndex, $currentMessages.Count - 1)
                            }
                        }
                    }
                }
            }
            default {
                $charPressed = $keyInfo.Character.ToString().ToUpper()

                if ($charPressed -eq 'A' -and $AllowActions) { # A - Select All
                    if ($currentMessages.Count -gt 0) {
                        foreach ($msg in $currentMessages) {
                            $spaceSelectedMessageIds.Add($msg.Id) | Out-Null
                        }
                    }
                } elseif ($charPressed -eq 'N' -and $AllowActions) { # N - Deselect All
                    $spaceSelectedMessageIds.Clear()
                } elseif ($charPressed -eq 'H' -and $AllowActions) { # H - Header Analysis
                    if ($currentMessages.Count -gt 0 -and $selectedEmailIndex -ge 0 -and $selectedEmailIndex -lt $currentMessages.Count) {
                        # Prepare display items for header analysis
                        $displayItems = @()
                        foreach ($msg in $currentMessages) {
                            $subject = if ($msg.Subject -and $msg.Subject.Length -gt 50) { $msg.Subject.Substring(0, 47) + "..." } elseif ($msg.Subject) { $msg.Subject } else { "(No Subject)" }
                            $sender = if ($msg.SenderEmailAddress -and $msg.SenderEmailAddress.Length -gt 30) { $msg.SenderEmailAddress.Substring(0, 27) + "..." } elseif ($msg.SenderEmailAddress) { $msg.SenderEmailAddress } else { "N/A" }

                            $displayItems += [PSCustomObject]@{
                                DisplayText = "$subject | From: $sender"
                                Message = $msg
                            }
                        }

                        # Show header analysis starting from selected email
                        Show-HeaderAnalysisView -UserEmail $UserEmail -AllMessages $displayItems -CurrentIndex $selectedEmailIndex
                    }
                } elseif ($charPressed -eq 'Q') {
                    $emailListLoopActive = $false
                }
            }
        }

        # Adjust indices
        if ($currentMessages -and $currentMessages.Count -gt 0) {
            $topDisplayIndex = [Math]::Max(0, [Math]::Min($topDisplayIndex, $currentMessages.Count - $currentDisplayLines))
            if ($topDisplayIndex -lt 0) {$topDisplayIndex = 0}
            $selectedEmailIndex = [Math]::Max(0, [Math]::Min($selectedEmailIndex, $currentMessages.Count - 1))
        } elseif (-not $currentMessages -or $currentMessages.Count -eq 0) {
            $emailListLoopActive = $false
        }
    }
}

# Function: Show-EmailDetails
function Show-EmailDetails {
    <#
    .SYNOPSIS
        Shows details of a single email
    .PARAMETER Message
        Message to display
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Message
    )

    Clear-Host
    Write-Host "`nEmail Details" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    Write-Host "Subject: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $(if ($Message.Subject) { $Message.Subject } else { "(No Subject)" }) -ForegroundColor $Global:ColorScheme.Value
    Write-Host ""

    Write-Host "From: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host "$($Message.SenderName) <$($Message.SenderEmailAddress)>" -ForegroundColor $Global:ColorScheme.Value
    Write-Host ""

    Write-Host "Date: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host (Format-SafeDateTime -DateTimeValue $Message.ReceivedDateTime) -ForegroundColor $Global:ColorScheme.Value
    Write-Host ""

    if ($Message.BodyPreview) {
        Write-Host "Preview:" -ForegroundColor $Global:ColorScheme.Label
        Write-Host $Message.BodyPreview -ForegroundColor $Global:ColorScheme.Normal
        Write-Host ""
    }

    Write-Host "Size: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    if ($Message.Size) {
        $sizeKB = [math]::Round($Message.Size / 1KB, 2)
        Write-Host "$sizeKB KB" -ForegroundColor $Global:ColorScheme.Value
    } else {
        Write-Host "N/A" -ForegroundColor $Global:ColorScheme.Value
    }
    Write-Host ""

    Write-Host "Has Attachments: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $(if ($Message.HasAttachments) { "Yes" } else { "No" }) -ForegroundColor $Global:ColorScheme.Value
    Write-Host ""

    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Read-Host "Press Enter to return"
}

# Export functions
Export-ModuleMember -Function Show-StandardizedEmailListView, Show-EmailDetails
