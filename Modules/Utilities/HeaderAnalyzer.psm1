<#
.SYNOPSIS
    Mail Header Analyzer module for MailCleanBuddy
.DESCRIPTION
    Analyzes email headers for debugging, security, and deliverability issues.
    Detects From/Reply-To mismatches, SPF/DKIM/DMARC status, and routing information.
#>

# Import dependencies

# Function: Get-EmailHeaders
function Get-EmailHeaders {
    <#
    .SYNOPSIS
        Retrieves full email headers
    .PARAMETER UserId
        User email address
    .PARAMETER MessageId
        Message ID
    .OUTPUTS
        Parsed header object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$MessageId
    )

    try {
        # Get message with internet headers
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId`?`$select=subject,from,replyTo,sender,internetMessageHeaders,receivedDateTime"
        $message = Invoke-MgGraphRequest -Method GET -Uri $uri

        return $message
    }
    catch {
        Write-Error "Error retrieving headers: $($_.Exception.Message)"
        return $null
    }
}

# Function: Analyze-EmailHeaders
function Analyze-EmailHeaders {
    <#
    .SYNOPSIS
        Analyzes email headers for issues
    .PARAMETER Message
        Message object with headers
    .OUTPUTS
        Analysis result object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Message
    )

    $analysis = [PSCustomObject]@{
        Warnings = @()
        Info = @()
        Security = @()
        Routing = @()
        FromReplyMismatch = $false
        SPFResult = "Unknown"
        DKIMResult = "Unknown"
        DMARCResult = "Unknown"
        MessagePath = @()
    }

    # Check From vs Reply-To mismatch
    if ($Message.from -and $Message.replyTo) {
        $fromEmail = $Message.from.emailAddress.address.ToLower()
        if ($Message.replyTo.Count -gt 0) {
            $replyToEmail = $Message.replyTo[0].emailAddress.address.ToLower()

            if ($fromEmail -ne $replyToEmail) {
                $analysis.FromReplyMismatch = $true
                $analysis.Warnings += "From/Reply-To mismatch detected"
                $analysis.Warnings += "  From: $fromEmail"
                $analysis.Warnings += "  Reply-To: $replyToEmail"
            }
        }
    }

    # Check sender vs from
    if ($Message.sender -and $Message.from) {
        $senderEmail = $Message.sender.emailAddress.address.ToLower()
        $fromEmail = $Message.from.emailAddress.address.ToLower()

        if ($senderEmail -ne $fromEmail) {
            $analysis.Info += "Sender differs from From (possible on-behalf-of sending)"
            $analysis.Info += "  Sender: $senderEmail"
            $analysis.Info += "  From: $fromEmail"
        }
    }

    # Analyze internet headers
    if ($Message.internetMessageHeaders) {
        foreach ($header in $Message.internetMessageHeaders) {
            $name = $header.name
            $value = $header.value

            # SPF Check
            if ($name -eq "Received-SPF") {
                if ($value -match "pass") {
                    $analysis.SPFResult = "Pass"
                    $analysis.Security += "✓ SPF: Pass"
                } elseif ($value -match "fail") {
                    $analysis.SPFResult = "Fail"
                    $analysis.Warnings += "⚠ SPF: Fail - Possible spoofing"
                } elseif ($value -match "softfail") {
                    $analysis.SPFResult = "SoftFail"
                    $analysis.Warnings += "⚠ SPF: SoftFail"
                } elseif ($value -match "neutral") {
                    $analysis.SPFResult = "Neutral"
                    $analysis.Info += "SPF: Neutral"
                }
            }

            # DKIM Check
            if ($name -eq "Authentication-Results") {
                if ($value -match "dkim=pass") {
                    $analysis.DKIMResult = "Pass"
                    $analysis.Security += "✓ DKIM: Pass"
                } elseif ($value -match "dkim=fail") {
                    $analysis.DKIMResult = "Fail"
                    $analysis.Warnings += "⚠ DKIM: Fail - Message may be tampered"
                } elseif ($value -match "dkim=none") {
                    $analysis.DKIMResult = "None"
                    $analysis.Info += "DKIM: Not signed"
                }

                # DMARC Check
                if ($value -match "dmarc=pass") {
                    $analysis.DMARCResult = "Pass"
                    $analysis.Security += "✓ DMARC: Pass"
                } elseif ($value -match "dmarc=fail") {
                    $analysis.DMARCResult = "Fail"
                    $analysis.Warnings += "⚠ DMARC: Fail"
                } elseif ($value -match "dmarc=none") {
                    $analysis.DMARCResult = "None"
                    $analysis.Info += "DMARC: No policy"
                }

                # SPF from Authentication-Results (more reliable)
                if ($value -match "spf=pass") {
                    $analysis.SPFResult = "Pass"
                } elseif ($value -match "spf=fail") {
                    $analysis.SPFResult = "Fail"
                }
            }

            # Message routing path
            if ($name -eq "Received") {
                $analysis.MessagePath += $value
            }

            # X-Mailer detection
            if ($name -eq "X-Mailer") {
                $analysis.Info += "Email Client: $value"
            }

            # Anti-spam headers
            if ($name -eq "X-Forefront-Antispam-Report" -or $name -eq "X-Microsoft-Antispam") {
                if ($value -match "SCL:([0-9])") {
                    $scl = $matches[1]
                    if ([int]$scl -ge 5) {
                        $analysis.Warnings += "⚠ High spam score detected (SCL: $scl)"
                    } else {
                        $analysis.Info += "Spam score (SCL): $scl"
                    }
                }
            }

            # Delivery errors
            if ($name -eq "X-Failed-Recipients") {
                $analysis.Warnings += "⚠ Failed recipients: $value"
            }
        }
    }

    return $analysis
}

# Function: Show-HeaderAnalyzer
function Show-HeaderAnalyzer {
    <#
    .SYNOPSIS
        Interactive header analyzer interface with arrow-key navigation
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        # Get cache to show email list
        $cache = Get-SenderCache

        if (-not $cache -or $cache.Count -eq 0) {
            Clear-Host
            $title = Get-LocalizedString "header_title" -FormatArgs @($UserEmail)
            Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
            Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""
            Write-Host "No cache data found. Please build the mailbox cache first." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Collect all messages from cache
        $allMessages = @()
        foreach ($domain in $cache.Keys) {
            if ($cache[$domain].Messages) {
                foreach ($msg in $cache[$domain].Messages) {
                    $msgId = if ($msg.MessageId) { $msg.MessageId } elseif ($msg.Id) { $msg.Id } else { $null }
                    if ($msgId) {
                        $allMessages += [PSCustomObject]@{
                            Id                 = $msgId
                            MessageId          = $msgId
                            Subject            = if ($msg.Subject) { $msg.Subject } else { "(No Subject)" }
                            SenderName         = if ($msg.SenderName) { $msg.SenderName } else { "N/A" }
                            SenderEmailAddress = if ($msg.SenderEmailAddress) { $msg.SenderEmailAddress } else { "N/A" }
                            ReceivedDateTime   = $msg.ReceivedDateTime
                        }
                    }
                }
            }
        }

        if ($allMessages.Count -eq 0) {
            Clear-Host
            $title = Get-LocalizedString "header_title" -FormatArgs @($UserEmail)
            Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
            Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""
            Write-Host "No messages found in cache." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Sort by received date (newest first)
        $allMessages = $allMessages | Sort-Object ReceivedDateTime -Descending

        # Prepare display items for selection
        $displayItems = @()
        foreach ($msg in $allMessages) {
            $subject = if ($msg.Subject.Length -gt 50) { $msg.Subject.Substring(0, 47) + "..." } else { $msg.Subject }
            $sender = if ($msg.SenderEmailAddress.Length -gt 30) { $msg.SenderEmailAddress.Substring(0, 27) + "..." } else { $msg.SenderEmailAddress }

            $displayItems += [PSCustomObject]@{
                DisplayText = "$subject | From: $sender"
                Message = $msg
            }
        }

        # Show selectable list
        $title = Get-LocalizedString "header_title" -FormatArgs @($UserEmail)
        $selected = Show-SelectableList -Title "$title - Select Email" -Items $displayItems -DisplayProperty "DisplayText" -PageSize 20

        if ($selected) {
            # Find index of selected message
            $selectedIndex = 0
            for ($i = 0; $i -lt $displayItems.Count; $i++) {
                if ($displayItems[$i].Message.Id -eq $selected.Message.Id) {
                    $selectedIndex = $i
                    break
                }
            }

            # Show header analysis with navigation
            Show-HeaderAnalysisView -UserEmail $UserEmail -AllMessages $displayItems -CurrentIndex $selectedIndex
        }
    }
    catch {
        Write-Error "Error in header analyzer: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Show-HeaderAnalysisView
function Show-HeaderAnalysisView {
    <#
    .SYNOPSIS
        Shows header analysis with prev/next navigation
    .PARAMETER UserEmail
        User email address
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
        [array]$AllMessages,

        [Parameter(Mandatory = $true)]
        [int]$CurrentIndex
    )

    try {
        $actionLoopActive = $true
        while ($actionLoopActive) {
            $currentItem = $AllMessages[$CurrentIndex]
            $messageId = $currentItem.Message.Id

            Clear-Host

            $title = Get-LocalizedString "header_title" -FormatArgs @($UserEmail)
            Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight

            # Show navigation info
            if ($AllMessages.Count -gt 0) {
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

            Write-Host (Get-LocalizedString "header_analyzing") -ForegroundColor $Global:ColorScheme.Info

            # Get headers
            $message = Get-EmailHeaders -UserId $UserEmail -MessageId $messageId

            if (-not $message) {
                Write-Host (Get-LocalizedString "header_notFound") -ForegroundColor $Global:ColorScheme.Error
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                return
            }

            # Analyze
            $analysis = Analyze-EmailHeaders -Message $message

            # Display results
            Clear-Host
            Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight

            # Show navigation info again
            if ($AllMessages.Count -gt 0) {
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

            # Message info
            Write-Host "📧 $(Get-LocalizedString 'header_messageInfo')" -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host "  Subject: $($message.subject)" -ForegroundColor $Global:ColorScheme.Normal
            Write-Host "  From: $($message.from.emailAddress.address)" -ForegroundColor $Global:ColorScheme.Normal
            if ($message.replyTo -and $message.replyTo.Count -gt 0) {
                Write-Host "  Reply-To: $($message.replyTo[0].emailAddress.address)" -ForegroundColor $Global:ColorScheme.Normal
            }
            Write-Host "  Received: $($message.receivedDateTime)" -ForegroundColor $Global:ColorScheme.Normal
            Write-Host ""

            # Security analysis
            Write-Host "🔒 $(Get-LocalizedString 'header_securityAnalysis')" -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host "  SPF: $($analysis.SPFResult)" -ForegroundColor $(if ($analysis.SPFResult -eq "Pass") { $Global:ColorScheme.Success } elseif ($analysis.SPFResult -eq "Fail") { $Global:ColorScheme.Error } else { $Global:ColorScheme.Warning })
            Write-Host "  DKIM: $($analysis.DKIMResult)" -ForegroundColor $(if ($analysis.DKIMResult -eq "Pass") { $Global:ColorScheme.Success } elseif ($analysis.DKIMResult -eq "Fail") { $Global:ColorScheme.Error } else { $Global:ColorScheme.Warning })
            Write-Host "  DMARC: $($analysis.DMARCResult)" -ForegroundColor $(if ($analysis.DMARCResult -eq "Pass") { $Global:ColorScheme.Success } elseif ($analysis.DMARCResult -eq "Fail") { $Global:ColorScheme.Error } else { $Global:ColorScheme.Warning })

            if ($analysis.Security.Count -gt 0) {
                Write-Host ""
                foreach ($sec in $analysis.Security) {
                    Write-Host "  $sec" -ForegroundColor $Global:ColorScheme.Success
                }
            }
            Write-Host ""

            # Warnings
            if ($analysis.Warnings.Count -gt 0) {
                Write-Host "⚠️  $(Get-LocalizedString 'header_warnings')" -ForegroundColor $Global:ColorScheme.SectionHeader
                Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border
                foreach ($warning in $analysis.Warnings) {
                    Write-Host "  $warning" -ForegroundColor $Global:ColorScheme.Warning
                }
                Write-Host ""
            }

            # Info
            if ($analysis.Info.Count -gt 0) {
                Write-Host "ℹ️  $(Get-LocalizedString 'header_info')" -ForegroundColor $Global:ColorScheme.SectionHeader
                Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border
                foreach ($info in $analysis.Info) {
                    Write-Host "  $info" -ForegroundColor $Global:ColorScheme.Info
                }
                Write-Host ""
            }

            # Routing path summary
            if ($analysis.MessagePath.Count -gt 0) {
                Write-Host "🌐 $(Get-LocalizedString 'header_routingPath')" -ForegroundColor $Global:ColorScheme.SectionHeader
                Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border
                Write-Host "  $(Get-LocalizedString 'header_hopCount' -FormatArgs @($analysis.MessagePath.Count))" -ForegroundColor $Global:ColorScheme.Info
                Write-Host ""
            }

            # Show available actions
            Write-Host "Available Actions:" -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  [R] Show Routing Details" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [E] Export Analysis" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [Q/Esc] Back" -ForegroundColor $Global:ColorScheme.Muted
            Write-Host ""

            # Read key
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

            switch ($keyInfo.VirtualKeyCode) {
                37 { # Left Arrow - Previous email
                    if ($CurrentIndex -gt 0) {
                        $CurrentIndex--
                    }
                }
                39 { # Right Arrow - Next email
                    if ($CurrentIndex -lt ($AllMessages.Count - 1)) {
                        $CurrentIndex++
                    }
                }
                27 { # Escape
                    $actionLoopActive = $false
                }
                default {
                    $charPressed = $keyInfo.Character.ToString().ToUpper()
                    if ($charPressed -eq 'Q') {
                        $actionLoopActive = $false
                    } elseif ($charPressed -eq 'R') {
                        # Show routing details
                        if ($analysis.MessagePath.Count -gt 0) {
                            Clear-Host
                            Write-Host "`n🌐 Routing Path Details" -ForegroundColor $Global:ColorScheme.SectionHeader
                            Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
                            Write-Host ""
                            $hopNum = 1
                            foreach ($hop in $analysis.MessagePath) {
                                Write-Host "Hop $hopNum :" -ForegroundColor $Global:ColorScheme.Label
                                Write-Host "  $hop" -ForegroundColor $Global:ColorScheme.Muted
                                Write-Host ""
                                $hopNum++
                            }
                            Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
                            Write-Host ""
                            Read-Host "Press Enter to continue"
                        }
                    } elseif ($charPressed -eq 'E') {
                        # Export analysis
                        Export-HeaderAnalysis -Message $message -Analysis $analysis
                        Read-Host "Press Enter to continue"
                    }
                }
            }
        }
    }
    catch {
        Write-Error "Error showing header analysis: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Export-HeaderAnalysis
function Export-HeaderAnalysis {
    <#
    .SYNOPSIS
        Exports header analysis to file
    .PARAMETER Message
        Message object
    .PARAMETER Analysis
        Analysis result
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Message,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Analysis
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\header_analysis.txt"
    $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    $report = @"
=== Email Header Analysis Report ===
Generated: $(Get-Date)

=== Message Info ===
Subject: $($Message.subject)
From: $($Message.from.emailAddress.address)
$(if ($Message.replyTo -and $Message.replyTo.Count -gt 0) { "Reply-To: $($Message.replyTo[0].emailAddress.address)" })
Received: $($Message.receivedDateTime)

=== Security Analysis ===
SPF: $($Analysis.SPFResult)
DKIM: $($Analysis.DKIMResult)
DMARC: $($Analysis.DMARCResult)
From/Reply-To Mismatch: $($Analysis.FromReplyMismatch)

=== Warnings ===
$($Analysis.Warnings -join "`n")

=== Info ===
$($Analysis.Info -join "`n")

=== Security Details ===
$($Analysis.Security -join "`n")

=== Message Routing Path ($($Analysis.MessagePath.Count) hops) ===
$($Analysis.MessagePath | ForEach-Object { $i = 1 } { "Hop $i : $_"; $i++ } | Out-String)

=== Full Internet Message Headers ===
$($Message.internetMessageHeaders | ForEach-Object { "$($_.name): $($_.value)" } | Out-String)
"@

    $report | Set-Content -Path $exportPath -Encoding UTF8

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
}

# Export functions
Export-ModuleMember -Function Show-HeaderAnalyzer, Show-HeaderAnalysisView, Analyze-EmailHeaders, Get-EmailHeaders
