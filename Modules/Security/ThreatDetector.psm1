<#
.SYNOPSIS
    Threat Detection module for MailCleanBuddy
.DESCRIPTION
    Detects phishing attempts, malware, spoofing, and other security threats in emails.
#>

# Import dependencies

# Threat database path
$script:ThreatDataPath = $null

# Known dangerous file extensions
$script:DangerousExtensions = @(
    '.exe', '.bat', '.cmd', '.com', '.pif', '.scr', '.vbs', '.js',
    '.jar', '.msi', '.dll', '.reg', '.ps1', '.hta', '.wsf', '.lnk'
)

# Phishing keywords (Dutch, English, German, French)
$script:PhishingKeywords = @(
    # Urgency
    'urgent', 'dringend', 'urgent', 'immédiat', 'immediately', 'onmiddellijk',
    'sofort', 'maintenant', 'action required', 'actie vereist', 'aktion erforderlich',
    # Account/Security
    'verify your account', 'verifieer je account', 'verifizieren Sie Ihr Konto',
    'vérifiez votre compte', 'suspended', 'opgeschort', 'gesperrt', 'suspendu',
    'blocked', 'geblokkeerd', 'blockiert', 'bloqué', 'unusual activity',
    'ongebruikelijke activiteit', 'ungewöhnliche Aktivität', 'activité inhabituelle',
    # Money/Payment
    'refund', 'terugbetaling', 'rückerstattung', 'remboursement', 'prize',
    'prijs', 'gewonnen', 'prix', 'click here', 'klik hier', 'klicken Sie hier',
    'cliquez ici', 'confirm identity', 'bevestig identiteit', 'bestätigen Sie Ihre Identität',
    'confirmez votre identité', 'reset password', 'wachtwoord resetten',
    'passwort zurücksetzen', 'réinitialiser le mot de passe'
)

# Function: Initialize-ThreatDetector
function Initialize-ThreatDetector {
    <#
    .SYNOPSIS
        Initializes threat detector database
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
    $script:ThreatDataPath = Join-Path $PSScriptRoot "..\..\threat_data_$sanitizedEmail.json"

    if (-not (Test-Path $script:ThreatDataPath)) {
        $initialData = @{
            DetectedThreats = @()
            QuarantinedEmails = @()
            WhitelistedSenders = @()
            LastScan = $null
        }
        $initialData | ConvertTo-Json -Depth 10 | Set-Content -Path $script:ThreatDataPath -Encoding UTF8
    }
}

# Function: Analyze-EmailThreat
function Analyze-EmailThreat {
    <#
    .SYNOPSIS
        Analyzes an email for threats
    .PARAMETER Message
        Email message to analyze
    .OUTPUTS
        Threat analysis object
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Message
    )

    # Handle both cache messages and Graph API messages
    $messageId = if ($Message.MessageId) { $Message.MessageId } elseif ($Message.Id) { $Message.Id } else { "" }
    $senderEmail = if ($Message.SenderEmailAddress) { $Message.SenderEmailAddress } elseif ($Message.From.EmailAddress.Address) { $Message.From.EmailAddress.Address } else { "" }
    $senderName = if ($Message.SenderName) { $Message.SenderName } elseif ($Message.From.EmailAddress.Name) { $Message.From.EmailAddress.Name } else { "" }
    $subject = if ($Message.Subject) { $Message.Subject } else { "" }
    $bodyPreview = if ($Message.BodyPreview) { $Message.BodyPreview } else { "" }

    $threat = [PSCustomObject]@{
        MessageId = $messageId
        Subject = $subject
        SenderEmail = $senderEmail
        SenderName = $senderName
        ThreatScore = 0
        ThreatLevel = "None"
        Indicators = @()
        ThreatTypes = @()
        IsWhitelisted = $false
        Timestamp = (Get-Date).ToString("o")
    }

    # Check whitelist first
    if (-not [string]::IsNullOrWhiteSpace($senderEmail) -and (Test-IsWhitelisted -SenderEmail $senderEmail)) {
        $threat.IsWhitelisted = $true
        return $threat
    }

    # Indicator 1: From/Reply-To mismatch (only for full Graph API messages)
    if ($Message.PSObject.Properties['from'] -and $Message.PSObject.Properties['replyTo'] -and
        $Message.from -and $Message.replyTo -and $Message.replyTo.Count -gt 0) {
        $fromEmail = $Message.from.emailAddress.address.ToLower()
        $replyToEmail = $Message.replyTo[0].emailAddress.address.ToLower()

        if ($fromEmail -ne $replyToEmail) {
            $threat.ThreatScore += 15
            $threat.Indicators += "From/Reply-To address mismatch detected"
            if ($threat.ThreatTypes -notcontains "Spoofing") {
                $threat.ThreatTypes += "Spoofing"
            }
        }
    }

    # Indicator 2: Display name impersonation
    if (-not [string]::IsNullOrWhiteSpace($senderName) -and -not [string]::IsNullOrWhiteSpace($senderEmail)) {
        $displayName = $senderName.ToLower()
        $emailAddress = $senderEmail.ToLower()

        # Check for common impersonation (e.g., "PayPal" from random domain)
        $trustedBrands = @('paypal', 'microsoft', 'google', 'amazon', 'apple', 'bank', 'ing', 'rabobank', 'abn amro')
        foreach ($brand in $trustedBrands) {
            if ($displayName -like "*$brand*" -and $emailAddress -notlike "*$brand*") {
                $threat.ThreatScore += 25
                $threat.Indicators += "Display name impersonation detected: '$displayName' from $emailAddress"
                if ($threat.ThreatTypes -notcontains "Phishing") {
                    $threat.ThreatTypes += "Phishing"
                }
                break
            }
        }
    }

    # Indicator 3: Phishing keywords in subject/body
    $contentToCheck = "$subject $bodyPreview".ToLower()
    $keywordMatches = 0

    foreach ($keyword in $script:PhishingKeywords) {
        if ($contentToCheck -like "*$keyword*") {
            $keywordMatches++
        }
    }

    if ($keywordMatches -ge 3) {
        $threat.ThreatScore += 20
        $threat.Indicators += "Multiple phishing keywords detected ($keywordMatches matches)"
        if ($threat.ThreatTypes -notcontains "Phishing") {
            $threat.ThreatTypes += "Phishing"
        }
    } elseif ($keywordMatches -ge 1) {
        $threat.ThreatScore += 10
        $threat.Indicators += "Phishing keywords detected ($keywordMatches matches)"
    }

    # Indicator 4: Suspicious links (shortened URLs)
    $suspiciousUrlPatterns = @('bit.ly', 'tinyurl', 'goo.gl', 't.co', 'ow.ly', 'is.gd')
    foreach ($pattern in $suspiciousUrlPatterns) {
        if ($contentToCheck -like "*$pattern*") {
            $threat.ThreatScore += 10
            $threat.Indicators += "Shortened URL detected: $pattern"
            if ($threat.ThreatTypes -notcontains "Phishing") {
                $threat.ThreatTypes += "Phishing"
            }
            break
        }
    }

    # Indicator 5: Suspicious sender domain
    if (-not [string]::IsNullOrWhiteSpace($senderEmail) -and $senderEmail -like "*@*") {
        $senderDomain = $senderEmail.Split('@')[1].ToLower()

        # Check for recently registered domains (simplified heuristic)
        if ($senderDomain -match '\d{4,}' -or $senderDomain -match '-[a-z]{20,}') {
            $threat.ThreatScore += 10
            $threat.Indicators += "Suspicious sender domain pattern: $senderDomain"
            if ($threat.ThreatTypes -notcontains "Spoofing") {
                $threat.ThreatTypes += "Spoofing"
            }
        }

        # Check for typosquatting
        $commonDomains = @('microsoft.com', 'google.com', 'paypal.com', 'amazon.com', 'apple.com')
        foreach ($trustedDomain in $commonDomains) {
            if ($senderDomain -ne $trustedDomain -and (Compare-StringSimilarity $senderDomain $trustedDomain) -gt 0.8) {
                $threat.ThreatScore += 30
                $threat.Indicators += "Possible typosquatting: $senderDomain similar to $trustedDomain"
                if ($threat.ThreatTypes -notcontains "Spoofing") {
                    $threat.ThreatTypes += "Spoofing"
                }
                break
            }
        }
    }

    # Indicator 6: Dangerous attachments
    if ($Message.HasAttachments) {
        # Note: We can't check actual extensions from cache, but we can flag as potential risk
        $threat.ThreatScore += 5
        $threat.Indicators += "Email contains attachments (potential risk)"

        # If we can access full message, check attachment extensions
        # This would require Graph API call - simplified for now
    }

    # Indicator 7: No subject or very generic subject
    if ([string]::IsNullOrWhiteSpace($subject) -or
        $subject -in @('Re:', 'Fwd:', 'Hello', 'Hi', 'Document', 'Invoice', 'Payment')) {
        $threat.ThreatScore += 5
        $threat.Indicators += "Generic or missing subject line"
    }

    # Determine threat level
    if ($threat.ThreatScore -ge 50) {
        $threat.ThreatLevel = "Critical"
    } elseif ($threat.ThreatScore -ge 30) {
        $threat.ThreatLevel = "High"
    } elseif ($threat.ThreatScore -ge 15) {
        $threat.ThreatLevel = "Medium"
    } elseif ($threat.ThreatScore -gt 0) {
        $threat.ThreatLevel = "Low"
    }

    return $threat
}

# Function: Compare-StringSimilarity
function Compare-StringSimilarity {
    <#
    .SYNOPSIS
        Compares two strings for similarity (Levenshtein distance based)
    .PARAMETER String1
        First string
    .PARAMETER String2
        Second string
    .OUTPUTS
        Similarity score (0-1)
    #>
    [CmdletBinding()]
    param(
        [string]$String1,
        [string]$String2
    )

    if ($String1 -eq $String2) { return 1.0 }
    if ([string]::IsNullOrWhiteSpace($String1) -or [string]::IsNullOrWhiteSpace($String2)) { return 0.0 }

    $maxLength = [Math]::Max($String1.Length, $String2.Length)
    $levenshteinDistance = Get-LevenshteinDistance $String1 $String2

    return 1.0 - ($levenshteinDistance / $maxLength)
}

# Function: Get-LevenshteinDistance
function Get-LevenshteinDistance {
    <#
    .SYNOPSIS
        Calculates Levenshtein distance between two strings
    #>
    [CmdletBinding()]
    param(
        [string]$String1,
        [string]$String2
    )

    if ($String1.Length -eq 0) { return $String2.Length }
    if ($String2.Length -eq 0) { return $String1.Length }

    # Use hashtable for 2D array simulation (more compatible)
    $d = @{}

    for ($i = 0; $i -le $String1.Length; $i++) {
        $d["$i,0"] = $i
    }
    for ($j = 0; $j -le $String2.Length; $j++) {
        $d["0,$j"] = $j
    }

    for ($i = 1; $i -le $String1.Length; $i++) {
        for ($j = 1; $j -le $String2.Length; $j++) {
            $cost = if ($String1[$i - 1] -eq $String2[$j - 1]) { 0 } else { 1 }

            $deletion = $d["$($i-1),$j"] + 1
            $insertion = $d["$i,$($j-1)"] + 1
            $substitution = $d["$($i-1),$($j-1)"] + $cost

            $d["$i,$j"] = [Math]::Min([Math]::Min($deletion, $insertion), $substitution)
        }
    }

    return $d["$($String1.Length),$($String2.Length)"]
}

# Function: Scan-MailboxForThreats
function Scan-MailboxForThreats {
    <#
    .SYNOPSIS
        Scans entire mailbox for threats
    .PARAMETER UserEmail
        User email address
    .OUTPUTS
        Array of threats found
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Write-Host ""
        Write-Host (Get-LocalizedString "threat_scanning") -ForegroundColor $Global:ColorScheme.Info

        $cache = Get-SenderCache
        $threats = @()

        if (-not $cache -or $cache.Count -eq 0) {
            Write-Warning "No cache data found. Please build the mailbox cache first."
            return @()
        }

        $totalEmails = 0
        foreach ($domain in $cache.Keys) {
            if ($cache[$domain].Messages) {
                $totalEmails += $cache[$domain].Messages.Count
            }
        }

        if ($totalEmails -eq 0) {
            Write-Warning "No messages found in cache."
            return @()
        }

        $processed = 0
        $progressId = 1

        foreach ($domain in $cache.Keys) {
            if (-not $cache[$domain].Messages) {
                continue
            }

            foreach ($message in $cache[$domain].Messages) {
                $processed++

                if ($processed % 50 -eq 0) {
                    Write-Progress -Id $progressId -Activity (Get-LocalizedString "threat_progressActivity") `
                        -Status (Get-LocalizedString "threat_progressStatus" -FormatArgs @($processed, $totalEmails)) `
                        -PercentComplete (($processed / $totalEmails) * 100)
                }

                $analysis = Analyze-EmailThreat -Message $message

                if ($analysis.ThreatScore -gt 0) {
                    $threats += $analysis
                }
            }
        }

        Write-Progress -Id $progressId -Activity (Get-LocalizedString "threat_progressActivity") -Completed

        # Save detected threats
        Save-DetectedThreats -Threats $threats

        return $threats | Sort-Object ThreatScore -Descending
    }
    catch {
        Write-Error "Error scanning for threats: $($_.Exception.Message)"
        return @()
    }
}

# Function: Show-ThreatDetector
function Show-ThreatDetector {
    <#
    .SYNOPSIS
        Interactive threat detector interface
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Initialize-ThreatDetector -UserEmail $UserEmail

        Clear-Host

        $title = Get-LocalizedString "threat_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host (Get-LocalizedString "threat_description") -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Scan for threats
        $threats = Scan-MailboxForThreats -UserEmail $UserEmail

        Write-Host ""
        Write-Host (Get-LocalizedString "threat_scanComplete") -ForegroundColor $Global:ColorScheme.Success
        Write-Host ""

        # Categorize threats
        $criticalThreats = $threats | Where-Object { $_.ThreatLevel -eq "Critical" }
        $highThreats = $threats | Where-Object { $_.ThreatLevel -eq "High" }
        $mediumThreats = $threats | Where-Object { $_.ThreatLevel -eq "Medium" }
        $lowThreats = $threats | Where-Object { $_.ThreatLevel -eq "Low" }

        # Display summary
        Write-Host (Get-LocalizedString "threat_summaryTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        if ($threats.Count -eq 0) {
            Write-Host "  ✅ $(Get-LocalizedString 'threat_noThreats')" -ForegroundColor $Global:ColorScheme.Success
        } else {
            Write-Host "  $(Get-LocalizedString 'threat_totalFound'): " -NoNewline
            Write-Host "$($threats.Count)" -ForegroundColor $Global:ColorScheme.Warning

            if ($criticalThreats.Count -gt 0) {
                Write-Host "  🔴 $(Get-LocalizedString 'threat_critical'): " -NoNewline
                Write-Host "$($criticalThreats.Count)" -ForegroundColor $Global:ColorScheme.Error
            }
            if ($highThreats.Count -gt 0) {
                Write-Host "  🟠 $(Get-LocalizedString 'threat_high'): " -NoNewline
                Write-Host "$($highThreats.Count)" -ForegroundColor $Global:ColorScheme.Warning
            }
            if ($mediumThreats.Count -gt 0) {
                Write-Host "  🟡 $(Get-LocalizedString 'threat_medium'): " -NoNewline
                Write-Host "$($mediumThreats.Count)" -ForegroundColor $Global:ColorScheme.Info
            }
            if ($lowThreats.Count -gt 0) {
                Write-Host "  🟢 $(Get-LocalizedString 'threat_low'): " -NoNewline
                Write-Host "$($lowThreats.Count)" -ForegroundColor $Global:ColorScheme.Normal
            }
        }

        Write-Host ""

        # Show top threats
        if ($threats.Count -gt 0) {
            Show-TopThreats -Threats ($threats | Select-Object -First 10)
        }

        # Menu
        Write-Host ""
        Write-Host (Get-LocalizedString "threat_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  1. $(Get-LocalizedString 'threat_viewAllThreats')" -ForegroundColor Green
        Write-Host "  2. $(Get-LocalizedString 'threat_quarantineSelected')" -ForegroundColor Yellow
        Write-Host "  3. $(Get-LocalizedString 'threat_manageWhitelist')" -ForegroundColor Cyan
        Write-Host "  4. $(Get-LocalizedString 'threat_exportReport')" -ForegroundColor Magenta
        Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
        Write-Host ""

        $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

        switch ($choice.ToUpper()) {
            "1" {
                if ($threats -and $threats.Count -gt 0) {
                    Show-AllThreats -Threats $threats
                } else {
                    Write-Host ""
                    Write-Host (Get-LocalizedString 'threat_noThreats') -ForegroundColor $Global:ColorScheme.Info
                }
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            }
            "2" {
                if ($threats -and $threats.Count -gt 0) {
                    Invoke-QuarantineThreats -UserEmail $UserEmail -Threats $threats
                } else {
                    Write-Host ""
                    Write-Host (Get-LocalizedString 'threat_noThreats') -ForegroundColor $Global:ColorScheme.Info
                }
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            }
            "3" {
                Manage-ThreatWhitelist
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            }
            "4" {
                if ($threats -and $threats.Count -gt 0) {
                    Export-ThreatReport -Threats $threats
                } else {
                    Write-Host ""
                    Write-Host (Get-LocalizedString 'threat_noThreats') -ForegroundColor $Global:ColorScheme.Info
                }
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            }
        }
    }
    catch {
        Write-Error "Error in threat detector: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Show-TopThreats
function Show-TopThreats {
    <#
    .SYNOPSIS
        Displays top threats
    .PARAMETER Threats
        Array of threats
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Threats
    )

    Write-Host (Get-LocalizedString "threat_topThreatsTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

    $format = "{0,-4} {1,-12} {2,10} {3,-35} {4,-30}"
    Write-Host ($format -f "#", "Level", "Score", "Subject", "Sender") -ForegroundColor $Global:ColorScheme.Header
    Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

    $index = 1
    foreach ($threat in $Threats) {
        $levelColor = switch ($threat.ThreatLevel) {
            "Critical" { $Global:ColorScheme.Error }
            "High" { $Global:ColorScheme.Warning }
            "Medium" { $Global:ColorScheme.Info }
            default { $Global:ColorScheme.Normal }
        }

        $subject = if ($threat.Subject.Length -gt 34) {
            $threat.Subject.Substring(0, 31) + "..."
        } else {
            $threat.Subject
        }

        $sender = if ($threat.SenderEmail.Length -gt 29) {
            $threat.SenderEmail.Substring(0, 26) + "..."
        } else {
            $threat.SenderEmail
        }

        Write-Host ($format -f $index, $threat.ThreatLevel, $threat.ThreatScore, $subject, $sender) -ForegroundColor $levelColor
        $index++
    }
}

# Function: Show-AllThreats
function Show-AllThreats {
    <#
    .SYNOPSIS
        Shows detailed view of all threats
    .PARAMETER Threats
        Array of threats
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Threats
    )

    Write-Host ""
    Write-Host (Get-LocalizedString "threat_detailsTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    foreach ($threat in $Threats) {
        $levelColor = switch ($threat.ThreatLevel) {
            "Critical" { $Global:ColorScheme.Error }
            "High" { $Global:ColorScheme.Warning }
            "Medium" { $Global:ColorScheme.Info }
            default { $Global:ColorScheme.Normal }
        }

        Write-Host "Subject: " -NoNewline
        Write-Host "$($threat.Subject)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "From: " -NoNewline
        Write-Host "$($threat.SenderName) <$($threat.SenderEmail)>" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "Threat Level: " -NoNewline
        Write-Host "$($threat.ThreatLevel) (Score: $($threat.ThreatScore))" -ForegroundColor $levelColor
        Write-Host "Threat Types: " -NoNewline
        Write-Host ($threat.ThreatTypes -join ", ") -ForegroundColor $Global:ColorScheme.Warning
        Write-Host "Indicators:"
        foreach ($indicator in $threat.Indicators) {
            Write-Host "  - $indicator" -ForegroundColor $Global:ColorScheme.Muted
        }
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
    }
}

# Function: Save-DetectedThreats
function Save-DetectedThreats {
    <#
    .SYNOPSIS
        Saves detected threats to database
    .PARAMETER Threats
        Array of threats
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Threats
    )

    try {
        $data = Get-Content -Path $script:ThreatDataPath -Raw | ConvertFrom-Json
        $data.DetectedThreats = $Threats
        $data.LastScan = (Get-Date).ToString("o")
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:ThreatDataPath -Encoding UTF8
    }
    catch {
        Write-Warning "Could not save detected threats: $($_.Exception.Message)"
    }
}

# Function: Test-IsWhitelisted
function Test-IsWhitelisted {
    <#
    .SYNOPSIS
        Checks if sender is whitelisted
    .PARAMETER SenderEmail
        Sender email address
    .OUTPUTS
        Boolean
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SenderEmail
    )

    try {
        $data = Get-Content -Path $script:ThreatDataPath -Raw | ConvertFrom-Json
        return $data.WhitelistedSenders -contains $SenderEmail.ToLower()
    }
    catch {
        return $false
    }
}

# Function: Invoke-QuarantineThreats
function Invoke-QuarantineThreats {
    <#
    .SYNOPSIS
        Quarantines selected threats
    .PARAMETER UserEmail
        User email address
    .PARAMETER Threats
        Array of threats
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Threats
    )

    $highPriorityThreats = $Threats | Where-Object { $_.ThreatLevel -in @("Critical", "High") }

    if ($highPriorityThreats.Count -eq 0) {
        Write-Host "`n$(Get-LocalizedString 'threat_noHighThreats')" -ForegroundColor $Global:ColorScheme.Info
        return
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "threat_quarantinePrompt" -FormatArgs @($highPriorityThreats.Count)) -ForegroundColor $Global:ColorScheme.Warning
    Write-Host ""

    $confirm = Show-Confirmation -Message (Get-LocalizedString "threat_confirmQuarantine")

    if ($confirm) {
        Write-Host ""
        Write-Host (Get-LocalizedString "threat_quarantining") -ForegroundColor $Global:ColorScheme.Info

        try {
            # Get or create Quarantine folder
            $quarantineFolder = $null
            $folders = Get-GraphMailFolders -UserId $UserId
            $quarantineFolder = $folders | Where-Object { $_.DisplayName -eq "Quarantine" } | Select-Object -First 1

            if (-not $quarantineFolder) {
                Write-Host "Creating Quarantine folder..." -ForegroundColor $Global:ColorScheme.Info
                $quarantineFolder = New-GraphMailFolder -UserId $UserId -DisplayName "Quarantine"
            }

            if (-not $quarantineFolder) {
                Write-Host "Failed to create Quarantine folder." -ForegroundColor $Global:ColorScheme.Error
                return
            }

            # Move high-priority threats to quarantine
            $movedCount = 0
            $failedCount = 0

            foreach ($threat in $highPriorityThreats) {
                try {
                    Move-GraphMessage -UserId $UserId -MessageId $threat.MessageId -DestinationFolderId $quarantineFolder.Id | Out-Null
                    $movedCount++
                    Write-Progress -Activity "Quarantining threats" -Status "Moved $movedCount of $($highPriorityThreats.Count)" `
                                   -PercentComplete (($movedCount / $highPriorityThreats.Count) * 100)
                } catch {
                    $failedCount++
                    Write-Verbose "Failed to quarantine message $($threat.MessageId): $($_.Exception.Message)"
                }
            }

            Write-Progress -Activity "Quarantining threats" -Completed

            Write-Host ""
            Write-Host "Quarantine complete:" -ForegroundColor $Global:ColorScheme.Success
            Write-Host "  Moved: $movedCount threat(s) to Quarantine folder" -ForegroundColor $Global:ColorScheme.Success
            if ($failedCount -gt 0) {
                Write-Host "  Failed: $failedCount threat(s)" -ForegroundColor $Global:ColorScheme.Warning
            }
        } catch {
            Write-Host "Error during quarantine: $($_.Exception.Message)" -ForegroundColor $Global:ColorScheme.Error
        }
    }
}

# Function: Manage-ThreatWhitelist
function Manage-ThreatWhitelist {
    <#
    .SYNOPSIS
        Manages threat detection whitelist
    #>
    [CmdletBinding()]
    param()

    $data = Get-Content -Path $script:ThreatDataPath -Raw | ConvertFrom-Json

    Write-Host ""
    Write-Host (Get-LocalizedString "threat_whitelistTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 80) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    if ($data.WhitelistedSenders.Count -eq 0) {
        Write-Host (Get-LocalizedString "threat_noWhitelisted") -ForegroundColor $Global:ColorScheme.Info
    } else {
        $index = 1
        foreach ($sender in $data.WhitelistedSenders) {
            Write-Host "  [$index] $sender" -ForegroundColor $Global:ColorScheme.Value
            $index++
        }
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "threat_whitelistOptions") -ForegroundColor $Global:ColorScheme.Info
    Write-Host "  [A] Add sender"
    Write-Host "  [R] Remove sender"
    Write-Host "  [Q] Back"
    Write-Host ""

    $choice = Read-Host "Choice"

    switch ($choice.ToUpper()) {
        "A" {
            $email = Read-Host "Enter email address to whitelist"
            if (-not [string]::IsNullOrWhiteSpace($email)) {
                $data.WhitelistedSenders += $email.ToLower()
                $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:ThreatDataPath -Encoding UTF8
                Write-Host "✓ Added to whitelist" -ForegroundColor $Global:ColorScheme.Success
            }
        }
        "R" {
            if ($data.WhitelistedSenders.Count -gt 0) {
                $num = Read-Host "Enter number to remove"
                if ($num -match '^\d+$' -and [int]$num -ge 1 -and [int]$num -le $data.WhitelistedSenders.Count) {
                    $removed = $data.WhitelistedSenders[[int]$num - 1]
                    $data.WhitelistedSenders = $data.WhitelistedSenders | Where-Object { $_ -ne $removed }
                    $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:ThreatDataPath -Encoding UTF8
                    Write-Host "✓ Removed from whitelist" -ForegroundColor $Global:ColorScheme.Success
                }
            }
        }
    }
}

# Function: Export-ThreatReport
function Export-ThreatReport {
    <#
    .SYNOPSIS
        Exports threat report to CSV
    .PARAMETER Threats
        Array of threats
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Threats
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\threat_report.csv"
    $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    $reportData = $Threats | Select-Object Subject, SenderEmail, SenderName, ThreatLevel, ThreatScore, @{
        Name = 'ThreatTypes'
        Expression = { $_.ThreatTypes -join '; ' }
    }, @{
        Name = 'Indicators'
        Expression = { $_.Indicators -join '; ' }
    }, Timestamp

    $reportData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
}

# Export functions
Export-ModuleMember -Function Show-ThreatDetector, Scan-MailboxForThreats, Analyze-EmailThreat, Initialize-ThreatDetector
