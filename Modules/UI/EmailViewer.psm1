<#
.SYNOPSIS
    Email Viewer module for MailCleanBuddy
.DESCRIPTION
    Provides email viewing, reading, and action functionality
#>

# Import dependencies

# Function: Show-EmailActionsMenu
function Show-EmailActionsMenu {
    <#
    .SYNOPSIS
        Shows email details and available actions
    .PARAMETER UserEmail
        User email address
    .PARAMETER MessageId
        Message ID to display
    .PARAMETER AllMessages
        Array of all messages for navigation
    .PARAMETER CurrentIndex
        Current message index in the array
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$MessageId,

        [Parameter(Mandatory = $false)]
        [array]$AllMessages = @(),

        [Parameter(Mandatory = $false)]
        [int]$CurrentIndex = -1
    )

    try {
        Clear-Host

        # Fetch email with full details
        $properties = "subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead"
        $message = Get-MgUserMessage -UserId $UserEmail -MessageId $MessageId -Property $properties -ErrorAction Stop

        if (-not $message) {
            Write-Host "`nEmail not found." -ForegroundColor $Global:ColorScheme.Error
            Read-Host "Press Enter to continue"
            return
        }

        $actionLoopActive = $true
        while ($actionLoopActive) {
            Clear-Host

            # Display email details with enhanced formatting
            Write-Host "`n📧 Email Details" -ForegroundColor $Global:ColorScheme.Highlight

            # Show navigation info if available
            if ($AllMessages.Count -gt 0 -and $CurrentIndex -ge 0) {
                $navInfo = "[$($CurrentIndex + 1) of $($AllMessages.Count)]"
                if ($CurrentIndex -gt 0) {
                    $navInfo += " [← Previous]"
                }
                if ($CurrentIndex -lt ($AllMessages.Count - 1)) {
                    $navInfo += " [Next →]"
                }
                Write-Host $navInfo -ForegroundColor $Global:ColorScheme.Muted
            }

            Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""

            # Subject with importance indicator
            $importanceIcon = if ($message.Importance -eq "high") { "⚠️ " } elseif ($message.Importance -eq "low") { "ℹ️ " } else { "" }
            Write-Host "📝 Subject      : " -NoNewline -ForegroundColor $Global:ColorScheme.Label
            Write-Host "$importanceIcon$(if ($message.Subject) { $message.Subject } else { '(No Subject)' })" -ForegroundColor $Global:ColorScheme.Value

            # From
            Write-Host "👤 From         : " -NoNewline -ForegroundColor $Global:ColorScheme.Label
            if ($message.From -and $message.From.EmailAddress) {
                Write-Host "$($message.From.EmailAddress.Name) <$($message.From.EmailAddress.Address)>" -ForegroundColor $Global:ColorScheme.Value
            } else {
                Write-Host "N/A" -ForegroundColor $Global:ColorScheme.Value
            }

            # To
            Write-Host "📨 To           : " -NoNewline -ForegroundColor $Global:ColorScheme.Label
            if ($message.ToRecipients) {
                $toList = ($message.ToRecipients | ForEach-Object { $_.EmailAddress.Address }) -join ", "
                Write-Host $toList -ForegroundColor $Global:ColorScheme.Value
            } else {
                Write-Host "N/A" -ForegroundColor $Global:ColorScheme.Value
            }

            # CC (if present)
            if ($message.CcRecipients -and $message.CcRecipients.Count -gt 0) {
                Write-Host "📋 CC           : " -NoNewline -ForegroundColor $Global:ColorScheme.Label
                $ccList = ($message.CcRecipients | ForEach-Object { $_.EmailAddress.Address }) -join ", "
                Write-Host $ccList -ForegroundColor $Global:ColorScheme.Value
            }

            # Date/Time
            Write-Host "🕐 Received     : " -NoNewline -ForegroundColor $Global:ColorScheme.Label
            try {
                Write-Host (Get-Date $message.ReceivedDateTime -Format "yyyy-MM-dd HH:mm:ss" -ErrorAction Stop) -ForegroundColor $Global:ColorScheme.Value
            } catch {
                Write-Host ($message.ReceivedDateTime.ToString()) -ForegroundColor $Global:ColorScheme.Value
            }

            # Attachments
            $attachIcon = if ($message.HasAttachments) { "📎 Yes" } else { "○ No" }
            Write-Host "📎 Attachments  : " -NoNewline -ForegroundColor $Global:ColorScheme.Label
            Write-Host $attachIcon -ForegroundColor $(if ($message.HasAttachments) { $Global:ColorScheme.Warning } else { $Global:ColorScheme.Value })

            # Read status
            $readIcon = if ($message.IsRead) { "✓ Read" } else { "○ Unread" }
            Write-Host "👁️  Status       : " -NoNewline -ForegroundColor $Global:ColorScheme.Label
            Write-Host $readIcon -ForegroundColor $(if ($message.IsRead) { $Global:ColorScheme.Success } else { $Global:ColorScheme.Warning })

            Write-Host ""
            Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""

            # Get and display full email body
            $bodyContent = ""
            $contentType = "text"

            if ($message.PSObject.Properties["Body"] -and $message.Body.PSObject.Properties["Content"]) {
                $bodyContent = $message.Body.Content
                $contentType = $message.Body.ContentType
            } elseif ($message.PSObject.Properties["body"] -and $message.body.content) {
                $bodyContent = $message.body.content
                $contentType = $message.body.contentType
            }

            # If body is not available in current message object, use preview
            if ([string]::IsNullOrWhiteSpace($bodyContent)) {
                if ($message.BodyPreview) {
                    $bodyContent = $message.BodyPreview
                } else {
                    $bodyContent = "(No body content available)"
                }
            }

            # Convert HTML to plain text if needed
            if ($contentType -eq "html") {
                $displayText = Convert-HtmlToPlainText -HtmlContent $bodyContent
            } else {
                $displayText = $bodyContent
            }

            # Limit body display to avoid extremely long emails
            $maxBodyLines = 30
            $bodyLines = $displayText -split "`r?`n"
            if ($bodyLines.Count -gt $maxBodyLines) {
                $displayText = ($bodyLines | Select-Object -First $maxBodyLines) -join "`n"
                $displayText += "`n`n... (Body truncated. Press [B] to view full body)"
            }

            Write-Host "📄 Body Preview:" -ForegroundColor $Global:ColorScheme.Label
            Write-Host $displayText -ForegroundColor $Global:ColorScheme.Normal

            Write-Host ""
            Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""

            # Show available actions with icons
            Write-Host "⚡ Available Actions:" -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  [B] 📖 View Full Body" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [O] 🌐 Open in Browser (HTML)" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [H] 🔒 Header Analysis (Security)" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [R] 📋 Raw Email Headers" -ForegroundColor $Global:ColorScheme.Info
            if ($message.HasAttachments) {
                Write-Host "  [D] 💾 Download Attachments" -ForegroundColor $Global:ColorScheme.Info
            }
            Write-Host "  [Del] 🗑️  Delete Email" -ForegroundColor $Global:ColorScheme.Warning
            Write-Host "  [V] 📁 Move to Folder" -ForegroundColor $Global:ColorScheme.Info
            if ($AllMessages.Count -gt 0 -and $CurrentIndex -ge 0) {
                Write-Host "  [←/→] Navigate to Previous/Next Email" -ForegroundColor $Global:ColorScheme.Info
            }
            Write-Host "  [Q/Esc] ⬅️  Back" -ForegroundColor $Global:ColorScheme.Muted
            Write-Host ""

            # Read key
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

            switch ($keyInfo.VirtualKeyCode) {
                46 { # Delete
                    $confirm = Show-Confirmation -Message "Delete this email?"
                    if ($confirm) {
                        try {
                            Remove-GraphMessage -UserId $UserEmail -MessageId $MessageId | Out-Null
                            Write-Host "`nEmail deleted successfully." -ForegroundColor $Global:ColorScheme.Success
                            Start-Sleep -Seconds 1
                            $actionLoopActive = $false
                        } catch {
                            Write-Host "`nFailed to delete email: $($_.Exception.Message)" -ForegroundColor $Global:ColorScheme.Error
                            Start-Sleep -Seconds 2
                        }
                    }
                }
                86 { # V - Move
                    $folder = Select-MailFolder -UserEmail $UserEmail
                    if ($folder) {
                        $confirm = Show-Confirmation -Message "Move this email to selected folder?"
                        if ($confirm) {
                            try {
                                Move-GraphMessage -UserId $UserEmail -MessageId $MessageId -DestinationFolderId $folder | Out-Null
                                Write-Host "`nEmail moved successfully." -ForegroundColor $Global:ColorScheme.Success
                                Start-Sleep -Seconds 1
                                $actionLoopActive = $false
                            } catch {
                                Write-Host "`nFailed to move email: $($_.Exception.Message)" -ForegroundColor $Global:ColorScheme.Error
                                Start-Sleep -Seconds 2
                            }
                        }
                    }
                }
                37 { # Left Arrow - Previous email
                    if ($AllMessages.Count -gt 0 -and $CurrentIndex -gt 0) {
                        $previousIndex = $CurrentIndex - 1
                        $previousMessage = $AllMessages[$previousIndex]

                        # Get the message ID
                        $previousMessageId = $null
                        if (-not [string]::IsNullOrWhiteSpace($previousMessage.Id)) {
                            $previousMessageId = $previousMessage.Id
                        } elseif (-not [string]::IsNullOrWhiteSpace($previousMessage.MessageId)) {
                            $previousMessageId = $previousMessage.MessageId
                        }

                        if (-not [string]::IsNullOrWhiteSpace($previousMessageId)) {
                            Show-EmailActionsMenu -UserEmail $UserEmail -MessageId $previousMessageId -AllMessages $AllMessages -CurrentIndex $previousIndex
                            $actionLoopActive = $false  # Exit current loop to show previous email
                        }
                    }
                }
                39 { # Right Arrow - Next email
                    if ($AllMessages.Count -gt 0 -and $CurrentIndex -ge 0 -and $CurrentIndex -lt ($AllMessages.Count - 1)) {
                        $nextIndex = $CurrentIndex + 1
                        $nextMessage = $AllMessages[$nextIndex]

                        # Get the message ID
                        $nextMessageId = $null
                        if (-not [string]::IsNullOrWhiteSpace($nextMessage.Id)) {
                            $nextMessageId = $nextMessage.Id
                        } elseif (-not [string]::IsNullOrWhiteSpace($nextMessage.MessageId)) {
                            $nextMessageId = $nextMessage.MessageId
                        }

                        if (-not [string]::IsNullOrWhiteSpace($nextMessageId)) {
                            Show-EmailActionsMenu -UserEmail $UserEmail -MessageId $nextMessageId -AllMessages $AllMessages -CurrentIndex $nextIndex
                            $actionLoopActive = $false  # Exit current loop to show next email
                        }
                    }
                }
                27 { $actionLoopActive = $false } # Escape
                default {
                    $charPressed = $keyInfo.Character.ToString().ToUpper()
                    if ($charPressed -eq 'B') { # View Body
                        Show-EmailBody -UserEmail $UserEmail -Message $message
                    } elseif ($charPressed -eq 'O') { # Open in Browser
                        Show-EmailInBrowser -UserEmail $UserEmail -MessageId $MessageId
                    } elseif ($charPressed -eq 'H') { # Header Analysis (Security)
                        # Prepare display items for header analysis navigation
                        $displayItems = @()
                        foreach ($msg in $AllMessages) {
                            # Get the message ID
                            $msgId = $null
                            if (-not [string]::IsNullOrWhiteSpace($msg.Id)) {
                                $msgId = $msg.Id
                            } elseif (-not [string]::IsNullOrWhiteSpace($msg.MessageId)) {
                                $msgId = $msg.MessageId
                            }

                            if (-not [string]::IsNullOrWhiteSpace($msgId)) {
                                $subject = if ($msg.Subject -and $msg.Subject.Length -gt 50) { $msg.Subject.Substring(0, 47) + "..." } elseif ($msg.Subject) { $msg.Subject } else { "(No Subject)" }
                                $sender = if ($msg.SenderEmailAddress -and $msg.SenderEmailAddress.Length -gt 30) { $msg.SenderEmailAddress.Substring(0, 27) + "..." } elseif ($msg.SenderEmailAddress) { $msg.SenderEmailAddress } else { "N/A" }

                                $displayItems += [PSCustomObject]@{
                                    DisplayText = "$subject | From: $sender"
                                    Message = $msg
                                }
                            }
                        }

                        if ($displayItems.Count -gt 0) {
                            Show-HeaderAnalysisView -UserEmail $UserEmail -AllMessages $displayItems -CurrentIndex $CurrentIndex
                        }
                    } elseif ($charPressed -eq 'R') { # Raw Headers
                        Show-EmailHeaders -UserEmail $UserEmail -MessageId $MessageId
                    } elseif ($charPressed -eq 'D' -and $message.HasAttachments) { # Download Attachments
                        Show-AttachmentDownloadMenu -UserEmail $UserEmail -MessageId $MessageId
                    } elseif ($charPressed -eq 'Q') {
                        $actionLoopActive = $false
                    }
                }
            }
        }
    }
    catch {
        Write-Error "Error showing email actions: $($_.Exception.Message)"
        Write-Host "`nAn error occurred." -ForegroundColor $Global:ColorScheme.Error
        Read-Host "Press Enter to continue"
    }
}

# Function: Show-EmailBody
function Show-EmailBody {
    <#
    .SYNOPSIS
        Shows the full body of an email
    .PARAMETER UserEmail
        User email address
    .PARAMETER Message
        Message object (can be partial)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Message
    )

    try {
        Clear-Host
        Write-Host "`n📖 Email Body - Full View" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host "📝 Subject: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host $(if ($Message.Subject) { $Message.Subject } else { "(No Subject)" }) -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""
        Write-Host "🕐 Received: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        try {
            Write-Host (Get-Date $Message.ReceivedDateTime -Format "yyyy-MM-dd HH:mm:ss" -ErrorAction Stop) -ForegroundColor $Global:ColorScheme.Value
        } catch {
            Write-Host ($Message.ReceivedDateTime.ToString()) -ForegroundColor $Global:ColorScheme.Value
        }
        Write-Host ""
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Get body content
        $bodyContent = ""
        $contentType = "text"

        if ($Message.PSObject.Properties["Body"] -and $Message.Body.PSObject.Properties["Content"]) {
            $bodyContent = $Message.Body.Content
            $contentType = $Message.Body.ContentType
        } elseif ($Message.PSObject.Properties["body"] -and $Message.body.content) {
            $bodyContent = $Message.body.content
            $contentType = $Message.body.contentType
        }

        # If body is not available, fetch it
        if ([string]::IsNullOrWhiteSpace($bodyContent)) {
            Write-Host "Fetching email body..." -ForegroundColor $Global:ColorScheme.Info
            try {
                $fullMessage = Get-MgUserMessage -UserId $UserEmail -MessageId $Message.Id -Property "body" -ErrorAction Stop
                if ($fullMessage -and $fullMessage.Body) {
                    $bodyContent = $fullMessage.Body.Content
                    $contentType = $fullMessage.Body.ContentType
                } else {
                    $bodyContent = $Message.BodyPreview
                    if ([string]::IsNullOrWhiteSpace($bodyContent)) {
                        $bodyContent = "(Could not fetch email body)"
                    }
                }
            } catch {
                Write-Host "Error fetching body: $($_.Exception.Message)" -ForegroundColor $Global:ColorScheme.Error
                $bodyContent = "(Error fetching email body)"
            }
        }

        # Convert HTML to plain text if needed
        if ($contentType -eq "html") {
            Write-Host "(Converting HTML to plain text...)" -ForegroundColor $Global:ColorScheme.Info
            Write-Host ""
            $displayText = Convert-HtmlToPlainText -HtmlContent $bodyContent
        } else {
            $displayText = $bodyContent
        }

        # Display the body
        Write-Host $displayText -ForegroundColor $Global:ColorScheme.Normal

        Write-Host ""
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""
        Write-Host "⬅️  Press Q or Esc to return" -ForegroundColor $Global:ColorScheme.Info

        # Wait for key
        while ($true) {
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)
            if ($keyInfo.VirtualKeyCode -eq 27 -or $keyInfo.Character.ToString().ToUpper() -eq 'Q') {
                break
            }
        }
    }
    catch {
        Write-Error "Error showing email body: $($_.Exception.Message)"
        Write-Host "`nAn error occurred." -ForegroundColor $Global:ColorScheme.Error
        Read-Host "Press Enter to continue"
    }
}

# Function: Show-AttachmentDownloadMenu
function Show-AttachmentDownloadMenu {
    <#
    .SYNOPSIS
        Shows menu to download attachments from an email
    .PARAMETER UserEmail
        User email address
    .PARAMETER MessageId
        Message ID
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$MessageId
    )

    try {
        Clear-Host
        Write-Host "`nDownload Attachments" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Get attachments
        $attachments = Get-MessageAttachments -UserId $UserEmail -MessageId $MessageId

        if (-not $attachments -or $attachments.Count -eq 0) {
            Write-Host "No attachments found." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        Write-Host "Found $($attachments.Count) attachment(s):" -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        $index = 1
        foreach ($att in $attachments) {
            Write-Host "  [$index] " -NoNewline -ForegroundColor $Global:ColorScheme.Muted
            Write-Host "$($att.Name)" -NoNewline -ForegroundColor $Global:ColorScheme.Value
            if ($att.Size) {
                $sizeKB = [math]::Round($att.Size / 1KB, 2)
                Write-Host " ($sizeKB KB)" -ForegroundColor $Global:ColorScheme.Muted
            } else {
                Write-Host ""
            }
            $index++
        }

        Write-Host ""
        $savePath = Read-Host "Enter save path (or press Enter for default: ./attachments)"
        if ([string]::IsNullOrWhiteSpace($savePath)) {
            $savePath = Join-Path $PSScriptRoot "..\..\attachments"
        }

        # Create directory if it doesn't exist
        Ensure-DirectoryExists -Path $savePath

        Write-Host ""
        Write-Host "Downloading attachments..." -ForegroundColor $Global:ColorScheme.Info

        $successCount = 0
        foreach ($att in $attachments) {
            try {
                Save-EmailAttachment -UserId $UserEmail -MessageId $MessageId -Attachment $att -SavePath $savePath
                Write-Host "  Downloaded: $($att.Name)" -ForegroundColor $Global:ColorScheme.Success
                $successCount++
            } catch {
                Write-Host "  Failed: $($att.Name) - $($_.Exception.Message)" -ForegroundColor $Global:ColorScheme.Error
            }
        }

        Write-Host ""
        Write-Host "$successCount of $($attachments.Count) attachment(s) downloaded successfully to: $savePath" -ForegroundColor $Global:ColorScheme.Success
        Read-Host "Press Enter to continue"
    }
    catch {
        Write-Error "Error downloading attachments: $($_.Exception.Message)"
        Write-Host "`nAn error occurred." -ForegroundColor $Global:ColorScheme.Error
        Read-Host "Press Enter to continue"
    }
}

# Function: Show-EmailHeaders
function Show-EmailHeaders {
    <#
    .SYNOPSIS
        Shows email transport headers
    .PARAMETER UserEmail
        User email address
    .PARAMETER MessageId
        Message ID
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$MessageId
    )

    try {
        Clear-Host
        Write-Host "`nEmail Headers" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host "Fetching email headers..." -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Fetch message with internet message headers
        $message = Get-MgUserMessage -UserId $UserEmail -MessageId $MessageId -Property "internetMessageHeaders,subject" -ErrorAction Stop

        if (-not $message) {
            Write-Host "Email not found." -ForegroundColor $Global:ColorScheme.Error
            Read-Host "Press Enter to continue"
            return
        }

        Write-Host "Subject: " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host $(if ($message.Subject) { $message.Subject } else { "(No Subject)" }) -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        if ($message.InternetMessageHeaders -and $message.InternetMessageHeaders.Count -gt 0) {
            Write-Host "Internet Message Headers:" -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ""

            foreach ($header in $message.InternetMessageHeaders) {
                Write-Host "$($header.Name): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
                Write-Host $header.Value -ForegroundColor $Global:ColorScheme.Normal
            }
        } else {
            Write-Host "No internet message headers available." -ForegroundColor $Global:ColorScheme.Warning
        }

        Write-Host ""
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""
        Write-Host "Press Q or Esc to return" -ForegroundColor $Global:ColorScheme.Info

        # Wait for key
        while ($true) {
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)
            if ($keyInfo.VirtualKeyCode -eq 27 -or $keyInfo.Character.ToString().ToUpper() -eq 'Q') {
                break
            }
        }
    }
    catch {
        Write-Error "Error showing email headers: $($_.Exception.Message)"
        Write-Host "`nAn error occurred." -ForegroundColor $Global:ColorScheme.Error
        Read-Host "Press Enter to continue"
    }
}

# Function: Show-EmailInBrowser
function Show-EmailInBrowser {
    <#
    .SYNOPSIS
        Opens email HTML content in default browser
    .PARAMETER UserEmail
        User email address
    .PARAMETER MessageId
        Message ID
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$MessageId
    )

    try {
        # Cleanup old temporary HTML files (older than 1 hour)
        try {
            $tempPath = [System.IO.Path]::GetTempPath()
            $oldTempFiles = Get-ChildItem -Path $tempPath -Filter "MailCleanBuddy_Email_*.html" -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-1) }

            foreach ($oldFile in $oldTempFiles) {
                try {
                    Remove-Item -Path $oldFile.FullName -Force -ErrorAction SilentlyContinue
                } catch {
                    # Silently ignore if file is in use
                }
            }
        } catch {
            # Silently ignore cleanup errors
        }

        Write-Host "`nFetching email content..." -ForegroundColor $Global:ColorScheme.Info

        # Fetch message with full body content
        $message = Get-MgUserMessage -UserId $UserEmail -MessageId $MessageId -Property "subject,body,from,receivedDateTime" -ErrorAction Stop

        if (-not $message) {
            Write-Host "Email not found." -ForegroundColor $Global:ColorScheme.Error
            Read-Host "Press Enter to continue"
            return
        }

        # Get HTML content
        $htmlContent = ""
        $contentType = "text"

        if ($message.Body -and $message.Body.Content) {
            $htmlContent = $message.Body.Content
            $contentType = $message.Body.ContentType
        }

        # Check if we have HTML content
        if ([string]::IsNullOrWhiteSpace($htmlContent)) {
            Write-Host "No email body content available." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return
        }

        # If content is plain text, wrap it in HTML
        if ($contentType -ne "html") {
            # Escape HTML special characters for plain text display
            $escapedContent = $htmlContent -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;'

            $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>$($message.Subject)</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background-color: #f5f5f5; }
        .email-container { background-color: white; padding: 30px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); max-width: 900px; margin: 0 auto; }
        .email-header { border-bottom: 2px solid #e0e0e0; padding-bottom: 15px; margin-bottom: 20px; }
        .email-subject { font-size: 24px; font-weight: bold; color: #333; margin-bottom: 10px; }
        .email-meta { font-size: 14px; color: #666; }
        .email-body { white-space: pre-wrap; font-family: 'Courier New', monospace; line-height: 1.6; }
    </style>
</head>
<body>
    <div class="email-container">
        <div class="email-header">
            <div class="email-subject">$($message.Subject)</div>
            <div class="email-meta">
                <strong>From:</strong> $($message.From.EmailAddress.Name) &lt;$($message.From.EmailAddress.Address)&gt;<br>
                <strong>Date:</strong> $($message.ReceivedDateTime)
            </div>
        </div>
        <div class="email-body">$escapedContent</div>
    </div>
</body>
</html>
"@
        } else {
            # For HTML content, wrap it with metadata header
            $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>$($message.Subject)</title>
    <style>
        .email-header-info { background-color: #f0f0f0; border: 1px solid #ddd; padding: 15px; margin-bottom: 20px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .email-header-info h3 { margin: 0 0 10px 0; color: #333; }
        .email-header-info p { margin: 5px 0; font-size: 14px; color: #555; }
    </style>
</head>
<body>
    <div class="email-header-info">
        <h3>$($message.Subject)</h3>
        <p><strong>From:</strong> $($message.From.EmailAddress.Name) &lt;$($message.From.EmailAddress.Address)&gt;</p>
        <p><strong>Date:</strong> $($message.ReceivedDateTime)</p>
    </div>
    $htmlContent
</body>
</html>
"@
        }

        # Create temp file
        $tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "MailCleanBuddy_Email_$($MessageId.Substring(0,8)).html")

        Write-Host "Creating temporary HTML file..." -ForegroundColor $Global:ColorScheme.Info
        Set-Content -Path $tempFile -Value $htmlContent -Encoding UTF8 -ErrorAction Stop

        Write-Host "Opening in default browser..." -ForegroundColor $Global:ColorScheme.Success
        Write-Host "File location: $tempFile" -ForegroundColor $Global:ColorScheme.Muted

        # Open in default browser
        Start-Process $tempFile

        Write-Host ""
        Write-Host "Email opened in browser." -ForegroundColor $Global:ColorScheme.Success
        Write-Host "Temporary file: $tempFile" -ForegroundColor $Global:ColorScheme.Muted
        Write-Host "Note: Temporary HTML files older than 1 hour are automatically cleaned up." -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""
        Read-Host "Press Enter to continue"
    }
    catch {
        Write-Error "Error opening email in browser: $($_.Exception.Message)"
        Write-Host "`nAn error occurred." -ForegroundColor $Global:ColorScheme.Error
        Read-Host "Press Enter to continue"
    }
}

# Export functions
Export-ModuleMember -Function Show-EmailActionsMenu, Show-EmailBody, Show-AttachmentDownloadMenu, Show-EmailHeaders, Show-EmailInBrowser
