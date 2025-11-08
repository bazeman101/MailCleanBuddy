<#
.SYNOPSIS
    Calendar Integration module for MailCleanBuddy
.DESCRIPTION
    Extracts calendar events from emails, detects meeting invitations, and syncs with calendar.
#>

# Import dependencies

# Calendar events database path
$script:CalendarDataPath = $null

# Meeting keywords (multilingual)
$script:MeetingKeywords = @(
    'meeting', 'vergadering', 'besprechung', 'réunion',
    'appointment', 'afspraak', 'termin', 'rendez-vous',
    'invitation', 'uitnodiging', 'einladung', 'calendrier',
    'schedule', 'agenda', 'zeitplan', 'webinar', 'conference',
    'call', 'gesprek', 'anruf', 'appel', 'zoom', 'teams', 'skype'
)

# Function: Initialize-CalendarSync
function Initialize-CalendarSync {
    <#
    .SYNOPSIS
        Initializes calendar sync database
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
    $script:CalendarDataPath = Join-Path $PSScriptRoot "..\..\calendar_events_$sanitizedEmail.json"

    if (-not (Test-Path $script:CalendarDataPath)) {
        $initialData = @{
            ExtractedEvents = @()
            SyncedEvents = @()
            LastSync = $null
        }
        $initialData | ConvertTo-Json -Depth 10 | Set-Content -Path $script:CalendarDataPath -Encoding UTF8
    }
}

# Function: Detect-CalendarEvents
function Detect-CalendarEvents {
    <#
    .SYNOPSIS
        Detects calendar events in emails
    .PARAMETER UserEmail
        User email address
    .OUTPUTS
        Array of detected events
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Write-Host ""
        Write-Host (Get-LocalizedString "calendar_detecting") -ForegroundColor $Global:ColorScheme.Info

        $cache = Get-SenderCache
        $events = @()

        $totalEmails = 0
        foreach ($domain in $cache.Keys) {
            $totalEmails += $cache[$domain].Messages.Count
        }

        $processed = 0
        $progressId = 1

        foreach ($domain in $cache.Keys) {
            foreach ($message in $cache[$domain].Messages) {
                $processed++

                if ($processed % 50 -eq 0) {
                    Write-Progress -Id $progressId -Activity (Get-LocalizedString "calendar_progressActivity") `
                        -Status (Get-LocalizedString "calendar_progressStatus" -FormatArgs @($processed, $totalEmails)) `
                        -PercentComplete (($processed / $totalEmails) * 100)
                }

                $event = Analyze-EmailForEvent -Message $message

                if ($event) {
                    $events += $event
                }
            }
        }

        Write-Progress -Id $progressId -Activity (Get-LocalizedString "calendar_progressActivity") -Completed

        return $events | Sort-Object { ConvertTo-SafeDateTime -DateTimeValue $_.StartDate }
    }
    catch {
        Write-Error "Error detecting calendar events: $($_.Exception.Message)"
        return @()
    }
}

# Function: Analyze-EmailForEvent
function Analyze-EmailForEvent {
    <#
    .SYNOPSIS
        Analyzes an email for calendar event indicators
    .PARAMETER Message
        Email message to analyze
    .OUTPUTS
        Event object if detected, $null otherwise
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Message
    )

    $score = 0
    $eventInfo = [PSCustomObject]@{
        MessageId = $Message.MessageId
        Subject = $Message.Subject
        SenderEmail = $Message.SenderEmailAddress
        SenderName = $Message.SenderName
        StartDate = $null
        EndDate = $null
        Location = $null
        IsOnlineMeeting = $false
        MeetingType = $null
        Confidence = 0
        ExtractedDate = $null
        HasIcsAttachment = $false
    }

    # Check 1: Subject contains meeting keywords
    $subjectLower = $Message.Subject.ToLower()
    foreach ($keyword in $script:MeetingKeywords) {
        if ($subjectLower -like "*$keyword*") {
            $score += 10
            break
        }
    }

    # Check 2: Has .ics attachment
    if ($Message.HasAttachments) {
        $score += 30
        $eventInfo.HasIcsAttachment = $true
    }

    # Check 3: Online meeting platforms
    $contentToCheck = "$($Message.Subject) $($Message.BodyPreview)".ToLower()

    $onlinePlatforms = @('zoom', 'teams', 'skype', 'webex', 'google meet', 'gotomeeting')
    foreach ($platform in $onlinePlatforms) {
        if ($contentToCheck -like "*$platform*") {
            $score += 15
            $eventInfo.IsOnlineMeeting = $true
            $eventInfo.MeetingType = $platform
            break
        }
    }

    # Check 4: Try to extract date/time
    $extractedDate = Extract-DateFromContent -Content $contentToCheck -ReceivedDate $Message.ReceivedDateTime

    if ($extractedDate) {
        $score += 20
        $eventInfo.StartDate = $extractedDate
        $eventInfo.ExtractedDate = $extractedDate
    }

    # Check 5: Location indicators
    $locationKeywords = @('location', 'locatie', 'standort', 'lieu', 'room', 'kamer', 'raum', 'salle')
    foreach ($keyword in $locationKeywords) {
        if ($contentToCheck -like "*$keyword*") {
            $score += 5
            break
        }
    }

    $eventInfo.Confidence = $score

    # Return event if confidence is high enough
    if ($score -ge 25) {
        return $eventInfo
    }

    return $null
}

# Function: Extract-DateFromContent
function Extract-DateFromContent {
    <#
    .SYNOPSIS
        Extracts date/time from email content
    .PARAMETER Content
        Email content to parse
    .PARAMETER ReceivedDate
        Email received date for context
    .OUTPUTS
        DateTime if found, $null otherwise
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Content,

        [Parameter(Mandatory = $false)]
        [string]$ReceivedDate
    )

    try {
        # Simple date pattern matching (can be improved with more sophisticated parsing)

        # Pattern 1: ISO date format (YYYY-MM-DD)
        if ($Content -match '\d{4}-\d{2}-\d{2}') {
            try {
                return [DateTime]::Parse($Matches[0])
            } catch {}
        }

        # Pattern 2: Common date formats (DD/MM/YYYY or MM/DD/YYYY)
        if ($Content -match '\d{1,2}[/-]\d{1,2}[/-]\d{4}') {
            try {
                return [DateTime]::Parse($Matches[0])
            } catch {}
        }

        # Pattern 3: Month name format (January 15, 2025)
        $monthNames = @('january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december',
                       'januari', 'februari', 'maart', 'mei', 'juni', 'juli', 'augustus', 'oktober', 'november', 'december')

        foreach ($month in $monthNames) {
            if ($Content -match "$month\s+\d{1,2}[,]?\s+\d{4}") {
                try {
                    return [DateTime]::Parse($Matches[0])
                } catch {}
            }
        }

        # Pattern 4: Relative dates (tomorrow, next week, etc.)
        if ($ReceivedDate) {
            $baseDate = [DateTime]::Parse($ReceivedDate)

            if ($Content -match 'tomorrow|morgen|demain') {
                return $baseDate.AddDays(1)
            }
            if ($Content -match 'next week|volgende week|nächste woche|semaine prochaine') {
                return $baseDate.AddDays(7)
            }
        }

        return $null
    }
    catch {
        return $null
    }
}

# Function: Show-CalendarSync
function Show-CalendarSync {
    <#
    .SYNOPSIS
        Interactive calendar sync interface
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Initialize-CalendarSync -UserEmail $UserEmail

        Clear-Host

        $title = Get-LocalizedString "calendar_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host (Get-LocalizedString "calendar_description") -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Detect events
        $events = Detect-CalendarEvents -UserEmail $UserEmail

        Write-Host ""
        Write-Host (Get-LocalizedString "calendar_detectionComplete") -ForegroundColor $Global:ColorScheme.Success
        Write-Host ""

        if ($events.Count -eq 0) {
            Write-Host (Get-LocalizedString "calendar_noEvents") -ForegroundColor $Global:ColorScheme.Info
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Display summary
        Write-Host (Get-LocalizedString "calendar_summaryTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""
        Write-Host "  $(Get-LocalizedString 'calendar_eventsFound'): " -NoNewline
        Write-Host "$($events.Count)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "  $(Get-LocalizedString 'calendar_withIcsAttachment'): " -NoNewline
        Write-Host "$(($events | Where-Object { $_.HasIcsAttachment }).Count)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "  $(Get-LocalizedString 'calendar_onlineMeetings'): " -NoNewline
        Write-Host "$(($events | Where-Object { $_.IsOnlineMeeting }).Count)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host ""

        # Show top events
        Show-TopEvents -Events ($events | Select-Object -First 15)

        # Menu
        Write-Host ""
        Write-Host (Get-LocalizedString "calendar_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  1. $(Get-LocalizedString 'calendar_viewAllEvents')" -ForegroundColor Green
        Write-Host "  2. $(Get-LocalizedString 'calendar_exportToIcs')" -ForegroundColor Cyan
        Write-Host "  3. $(Get-LocalizedString 'calendar_exportReport')" -ForegroundColor Magenta
        Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
        Write-Host ""

        $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

        switch ($choice.ToUpper()) {
            "1" {
                Show-AllEvents -Events $events
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            }
            "2" {
                Export-EventsToIcs -Events $events
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            }
            "3" {
                Export-CalendarReport -Events $events
                Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            }
        }

        # Save extracted events
        Save-ExtractedEvents -Events $events
    }
    catch {
        Write-Error "Error in calendar sync: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Show-TopEvents
function Show-TopEvents {
    <#
    .SYNOPSIS
        Displays top detected events
    .PARAMETER Events
        Array of events
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Events
    )

    Write-Host (Get-LocalizedString "calendar_topEventsTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

    $format = "{0,-4} {1,-15} {2,12} {3,-40} {4,-20}"
    Write-Host ($format -f "#", "Date", "Confidence", "Subject", "Type") -ForegroundColor $Global:ColorScheme.Header
    Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border

    $index = 1
    foreach ($event in $Events) {
        $date = if ($event.StartDate) {
            (ConvertTo-SafeDateTime -DateTimeValue $event.StartDate).ToString("yyyy-MM-dd HH:mm")
        } else {
            "Not detected"
        }

        $subject = if ($event.Subject.Length -gt 39) {
            $event.Subject.Substring(0, 36) + "..."
        } else {
            $event.Subject
        }

        $type = if ($event.HasIcsAttachment) {
            "ICS Invite"
        } elseif ($event.IsOnlineMeeting) {
            $event.MeetingType
        } else {
            "Detected"
        }

        $confidenceColor = if ($event.Confidence -ge 50) {
            $Global:ColorScheme.Success
        } elseif ($event.Confidence -ge 30) {
            $Global:ColorScheme.Info
        } else {
            $Global:ColorScheme.Warning
        }

        Write-Host ($format -f $index, $date, "$($event.Confidence)%", $subject, $type) -ForegroundColor $confidenceColor
        $index++
    }
}

# Function: Show-AllEvents
function Show-AllEvents {
    <#
    .SYNOPSIS
        Shows detailed view of all events
    .PARAMETER Events
        Array of events
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Events
    )

    Write-Host ""
    Write-Host (Get-LocalizedString "calendar_eventDetailsTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    foreach ($event in $Events) {
        Write-Host "Subject: " -NoNewline
        Write-Host "$($event.Subject)" -ForegroundColor $Global:ColorScheme.Value
        Write-Host "From: " -NoNewline
        Write-Host "$($event.SenderName) <$($event.SenderEmail)>" -ForegroundColor $Global:ColorScheme.Value

        if ($event.StartDate) {
            Write-Host "Date: " -NoNewline
            Write-Host "$($event.StartDate)" -ForegroundColor $Global:ColorScheme.Value
        }

        Write-Host "Type: " -NoNewline
        if ($event.HasIcsAttachment) {
            Write-Host "Calendar Invitation (ICS)" -ForegroundColor $Global:ColorScheme.Info
        } elseif ($event.IsOnlineMeeting) {
            Write-Host "Online Meeting ($($event.MeetingType))" -ForegroundColor $Global:ColorScheme.Info
        } else {
            Write-Host "Detected Event" -ForegroundColor $Global:ColorScheme.Normal
        }

        Write-Host "Confidence: " -NoNewline
        Write-Host "$($event.Confidence)%" -ForegroundColor $Global:ColorScheme.Value

        Write-Host ("-" * 100) -ForegroundColor $Global:ColorScheme.Border
    }
}

# Function: Export-EventsToIcs
function Export-EventsToIcs {
    <#
    .SYNOPSIS
        Exports events to ICS (iCalendar) format
    .PARAMETER Events
        Array of events
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Events
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\calendar_events.ics"
    $exportPath = Read-Host (Get-LocalizedString "calendar_exportPathIcs" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    try {
        $icsContent = "BEGIN:VCALENDAR`r`n"
        $icsContent += "VERSION:2.0`r`n"
        $icsContent += "PRODID:-//MailCleanBuddy//Calendar Sync//EN`r`n"
        $icsContent += "CALSCALE:GREGORIAN`r`n"

        foreach ($event in $Events) {
            if ($event.StartDate) {
                $icsContent += "BEGIN:VEVENT`r`n"
                $icsContent += "UID:$($event.MessageId)@mailcleanbuddy`r`n"

                $startDate = (ConvertTo-SafeDateTime -DateTimeValue $event.StartDate).ToString("yyyyMMdd\THHmmss")
                $icsContent += "DTSTART:$startDate`r`n"

                # Default 1 hour duration if no end date
                if ($event.EndDate) {
                    $endDate = (ConvertTo-SafeDateTime -DateTimeValue $event.EndDate).ToString("yyyyMMdd\THHmmss")
                } else {
                    $endDate = (ConvertTo-SafeDateTime -DateTimeValue $event.StartDate).AddHours(1).ToString("yyyyMMdd\THHmmss")
                }
                $icsContent += "DTEND:$endDate`r`n"

                $icsContent += "SUMMARY:$($event.Subject -replace ',', '\,')`r`n"
                $icsContent += "ORGANIZER:MAILTO:$($event.SenderEmail)`r`n"

                if ($event.Location) {
                    $icsContent += "LOCATION:$($event.Location)`r`n"
                }

                if ($event.IsOnlineMeeting) {
                    $icsContent += "DESCRIPTION:Online meeting via $($event.MeetingType)`r`n"
                }

                $icsContent += "END:VEVENT`r`n"
            }
        }

        $icsContent += "END:VCALENDAR`r`n"

        $icsContent | Set-Content -Path $exportPath -Encoding UTF8

        Write-Host ""
        Write-Host (Get-LocalizedString "calendar_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
    }
    catch {
        Write-Error "Error exporting to ICS: $($_.Exception.Message)"
    }
}

# Function: Export-CalendarReport
function Export-CalendarReport {
    <#
    .SYNOPSIS
        Exports calendar report to CSV
    .PARAMETER Events
        Array of events
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Events
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\calendar_report.csv"
    $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    $Events | Select-Object Subject, SenderName, SenderEmail, StartDate, EndDate, IsOnlineMeeting, MeetingType, HasIcsAttachment, Confidence |
        Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
}

# Function: Save-ExtractedEvents
function Save-ExtractedEvents {
    <#
    .SYNOPSIS
        Saves extracted events to database
    .PARAMETER Events
        Array of events
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Events
    )

    try {
        $data = Get-Content -Path $script:CalendarDataPath -Raw | ConvertFrom-Json
        $data.ExtractedEvents = $Events
        $data.LastSync = (Get-Date).ToString("o")
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:CalendarDataPath -Encoding UTF8
    }
    catch {
        Write-Warning "Could not save extracted events: $($_.Exception.Message)"
    }
}

# Export functions
Export-ModuleMember -Function Show-CalendarSync, Detect-CalendarEvents, Initialize-CalendarSync
