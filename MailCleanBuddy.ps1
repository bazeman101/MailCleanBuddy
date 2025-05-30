<#
.SYNOPSIS
    Provides interactive menu-driven mailbox management for Microsoft 365.
.DESCRIPTION
    This script connects to Microsoft 365 using interactive login and
    offers a menu to perform various mailbox operations like indexing,
    managing emails by sender, and searching emails.
.PARAMETER MailboxEmail
    The email address of the mailbox to manage.
.EXAMPLE
    .\MailCleanBuddy.ps1 -MailboxEmail "user@example.com"
    This command will connect to "user@example.com" and display the main menu.
.NOTES
    Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Mail modules.
    The script will attempt to install them if not found.
    Ensure you have the necessary permissions (Microsoft Graph: Mail.Read, Mail.ReadWrite) to access the specified mailbox.
    Mail.ReadWrite is required for deleting or moving emails.
    You will be prompted to consent to these permissions on first run.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$MailboxEmail,

    [Parameter(Mandatory = $false)]
    [switch]$TestRun,

    [Parameter(Mandatory = $false)]
    [int]$MaxEmailsToIndex = 0 # Nieuwe parameter om het max aantal te indexeren mails te specificeren
)

# Set console window size
$desiredHeight = 55 # Minimaal 50 regels + wat marge
$desiredWidth = 150 # Voldoende breedte
try {
    # Controleer of de UI interactief is voordat we de venstergrootte proberen aan te passen
    if ($Host.UI.GetType().Name -notmatch "ConsoleHostUserInterface") {
        Write-Verbose "Niet-interactieve host gedetecteerd, console grootte wordt niet aangepast."
    } else {
        $currentWindowSize = $Host.UI.RawUI.WindowSize
        $bufferSize = $Host.UI.RawUI.BufferSize

        $newWidth = $desiredWidth
        $newHeight = $desiredHeight

        # Buffer width must be >= window width
        if ($bufferSize.Width -lt $newWidth) {
            $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size ($newWidth, $bufferSize.Height)
        }
        # Buffer height must be >= window height
        # Als we de bufferhoogte vergroten, moeten we mogelijk de huidige bufferhoogte gebruiken als die al groter is dan de gewenste vensterhoogte
        $newBufferHeight = [Math]::Max($bufferSize.Height, $newHeight)
        if ($Host.UI.RawUI.BufferSize.Width -lt $newWidth -or $Host.UI.RawUI.BufferSize.Height -lt $newBufferHeight) {
             $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size ([Math]::Max($Host.UI.RawUI.BufferSize.Width, $newWidth), $newBufferHeight)
        }

        $Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size ($newWidth, $newHeight)

        # Zorg ervoor dat de buffer minstens zo groot is als het venster.
        $finalBufferWidth = [Math]::Max($Host.UI.RawUI.BufferSize.Width, $newWidth)
        $finalBufferHeight = [Math]::Max($Host.UI.RawUI.BufferSize.Height, $newHeight)
        if ($Host.UI.RawUI.BufferSize.Width -lt $finalBufferWidth -or $Host.UI.RawUI.BufferSize.Height -lt $finalBufferHeight) {
            $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size ($finalBufferWidth, $finalBufferHeight)
        }
        Write-Verbose "Console size set to Width: $newWidth, Height: $newHeight"
    }
} catch {
    Write-Warning "Could not set console window size: $($_.Exception.Message)"
}

<#
.PARAMETER MaxEmailsToIndex
    Specificeert het maximale aantal nieuwste e-mails dat moet worden geïndexeerd.
    Als deze waarde groter is dan 0, overschrijft dit de -TestRun switch voor het aantal te indexeren mails.
    Standaard (0) wordt de -TestRun logica (100 mails) of een volledige indexering (alle mails) gebruikt.
#>

# Script-level cache for sender information
$Script:SenderCache = $null
$Script:CacheFilePath = $null

# Functie om het pad naar het cachebestand te bepalen
function Get-CacheFilePath {
    param (
        [string]$MailboxEmail
    )
    $safeEmail = $MailboxEmail -replace "[^a-zA-Z0-9.-_]", "_"
    $cacheFileName = "mailcleanbuddy_cache_$($safeEmail).json"
    $Script:CacheFilePath = Join-Path -Path $PSScriptRoot -ChildPath $cacheFileName
    Write-Verbose "Cache file path set to: $($Script:CacheFilePath)"
}

# Functie om de lokale cache te laden
function Load-LocalCache {
    if (-not $Script:CacheFilePath) {
        Write-Warning "Cache file path is not set. Cannot load cache."
        return $false
    }
    if (Test-Path $Script:CacheFilePath) {
        try {
            Write-Host "Lokale cache gevonden. Bezig met laden: $($Script:CacheFilePath)"
            Write-Progress -Activity "Cache Laden" -Status "Bezig met laden van lokale cache..." -PercentComplete 30

            $jsonContent = Get-Content -Path $Script:CacheFilePath -Raw -ErrorAction Stop
            Write-Progress -Activity "Cache Laden" -Status "Cachebestand gelezen, converteren..." -PercentComplete 60
            $loadedCache = ConvertFrom-Json -InputObject $jsonContent -ErrorAction Stop

            Write-Progress -Activity "Cache Laden" -Status "Cache geconverteerd, berichten verwerken..." -PercentComplete 80
            # Converteer Messages array terug naar List<PSObject> voor elke entry
            $tempCache = @{}
            foreach ($key in $loadedCache.PSObject.Properties.Name) {
                $entry = $loadedCache.$key
                if ($entry.Messages -is [System.Array]) {
                    $messageList = [System.Collections.Generic.List[PSObject]]::new()
                    foreach ($msg in $entry.Messages) {
                        $messageList.Add($msg)
                    }
                    $entry.Messages = $messageList
                }
                $tempCache[$key] = $entry
            }
            $Script:SenderCache = $tempCache
            Write-Host "Lokale cache succesvol geladen."
            Write-Progress -Activity "Cache Laden" -Status "Cache succesvol geladen." -Completed
            return $true
        } catch {
            Write-Warning "Fout bij het laden of parsen van de lokale cache: $($_.Exception.Message). De cache wordt genegeerd en een volledige indexering zal worden uitgevoerd."
            $Script:SenderCache = $null # Zorg ervoor dat de cache leeg is bij een fout
            Write-Progress -Activity "Cache Laden" -Status "Fout bij laden van cache." -Completed
            return $false
        }
    } else {
        Write-Host "Geen lokale cache gevonden op: $($Script:CacheFilePath)"
        return $false
    }
}

# Functie om de lokale cache op te slaan
function Save-LocalCache {
    if (-not $Script:CacheFilePath) {
        Write-Warning "Cache file path is not set. Cannot save cache."
        return
    }
    if ($null -eq $Script:SenderCache) {
        Write-Warning "SenderCache is leeg. Cache wordt niet opgeslagen."
        return
    }
    try {
        Write-Host "Lokale cache opslaan naar: $($Script:CacheFilePath)"
        # Gebruik een voldoende hoge diepte voor ConvertTo-Json
        $jsonContent = $Script:SenderCache | ConvertTo-Json -Depth 10 -ErrorAction Stop
        Set-Content -Path $Script:CacheFilePath -Value $jsonContent -ErrorAction Stop
        Write-Host "Lokale cache succesvol opgeslagen."
    } catch {
        Write-Warning "Fout bij het opslaan van de lokale cache: $($_.Exception.Message)"
    }
}

# Placeholder functions for menu items
function Index-Mailbox {
    param($UserId)

    Write-Host "Starten met indexeren van mailbox voor $UserId..."
    if ($MaxEmailsToIndex -gt 0) {
        Write-Warning "** MaxEmailsToIndex ACTIEF: Maximaal de laatste $MaxEmailsToIndex e-mails worden geïndexeerd. **"
    } elseif ($TestRun.IsPresent) {
        Write-Warning "** TESTMODUS ACTIEF: Alleen de laatste 100 e-mails worden geïndexeerd. **"
    } # Anders, volledige indexering (geen specifieke waarschuwing hier nodig)

    $Script:SenderCache = @{} # Reset of initialiseer de cache

    try {
        $baseMessageProperties = "id,subject,sender,receivedDateTime,toRecipients,categories" # Verwijder 'size' en 'hasAttachments'
        $messageSizeMapiPropertyId = "Integer 0x0E08" # PR_MESSAGE_SIZE
        $messageHasAttachMapiPropertyId = "Boolean 0x0E1B" # PR_HASATTACH
        $expandExtendedProperties = "singleValueExtendedProperties(`$filter=id eq '$messageSizeMapiPropertyId' or id eq '$messageHasAttachMapiPropertyId')"
        $messages = $null

        # Bouw de parameters voor Get-MgUserMessage
        $getMgUserMessageParams = @{
            UserId         = $UserId
            Property       = $baseMessageProperties
            ExpandProperty = $expandExtendedProperties
            ErrorAction    = "Stop"
        }

        if ($MaxEmailsToIndex -gt 0) {
            $getMgUserMessageParams.Top = $MaxEmailsToIndex
            $getMgUserMessageParams.OrderBy = "receivedDateTime desc"
            Write-Host "Configuratie: Ophalen van de laatste $MaxEmailsToIndex berichten (incl. MAPI size)."
        } elseif ($TestRun.IsPresent) {
            $getMgUserMessageParams.Top = 100
            $getMgUserMessageParams.OrderBy = "receivedDateTime desc"
            Write-Host "Configuratie: Ophalen van de laatste 100 berichten (Testmodus, incl. MAPI size)."
        } else {
            $getMgUserMessageParams.All = $true
            Write-Host "Configuratie: Ophalen van alle berichten (Volledige modus, incl. MAPI size). Dit kan enige tijd duren."
        }

        Write-Host "Berichten ophalen..."
        $messages = Get-MgUserMessage @getMgUserMessageParams
        Write-Host "Berichten succesvol opgehaald."

        if ($null -eq $messages -or $messages.Count -eq 0) {
            Write-Warning "Geen berichten gevonden in de mailbox tijdens het indexeren."
            # Read-Host "Druk op Enter om terug te keren naar het hoofdmenu" # Verwijderd
            return
        }

        Write-Host "$($messages.Count) berichten gevonden. Verwerken van afzenders..."

        $processedCount = 0
        $totalMessages = $messages.Count
        $updateInterval = [math]::Ceiling($totalMessages / 20) # Update progress approximately 20 times
        if ($updateInterval -eq 0) {$updateInterval = 1}


        foreach ($message in $messages) {
            $processedCount++
            if ($processedCount % $updateInterval -eq 0 -or $processedCount -eq $totalMessages) {
                Write-Progress -Activity "Mailbox Indexeren" -Status "Verwerken van berichten..." -PercentComplete (($processedCount / $totalMessages) * 100) -CurrentOperation "$processedCount van $totalMessages berichten verwerkt."
            }

            $emailSenderAddressInfo = $message.Sender.EmailAddress
            if ($emailSenderAddressInfo -and $emailSenderAddressInfo.Address) {
                # Groepeer op domein
                $senderFullAddress = $emailSenderAddressInfo.Address
                $domain = ($senderFullAddress -split '@')[1]
                if ([string]::IsNullOrWhiteSpace($domain)) {
                    $domain = "onbekend_domein" # Fallback voor ongeldige e-mailadressen
                }
                $domainKey = $domain.ToLowerInvariant()
                # De 'naam' voor de cache entry wordt nu het domein zelf.
                # De oorspronkelijke $senderName ($sender.Name) wordt niet meer direct gebruikt voor de groepering.

                # Bepaal de grootte van het bericht
                # Bepaal de grootte van het bericht
                $currentMessageSize = $null
                $mapiSizeProp = $message.SingleValueExtendedProperties | Where-Object { $_.Id -eq $messageSizeMapiPropertyId } | Select-Object -First 1
                if ($mapiSizeProp -and $mapiSizeProp.Value) {
                    try { $currentMessageSize = [long]$mapiSizeProp.Value } catch { Write-Verbose "Kon MAPI size '$($mapiSizeProp.Value)' niet converteren (Index) ID $($message.Id)." }
                } # Fallback naar $message.size is niet meer nodig/mogelijk

                # Bepaal of er bijlagen zijn
                $currentHasAttachments = $false # Default
                $mapiAttachProp = $message.SingleValueExtendedProperties | Where-Object { $_.Id -eq $messageHasAttachMapiPropertyId } | Select-Object -First 1
                if ($mapiAttachProp -and $mapiAttachProp.Value -ne $null) {
                    try { $currentHasAttachments = [System.Convert]::ToBoolean($mapiAttachProp.Value) } catch { Write-Verbose "Kon MAPI hasAttach '$($mapiAttachProp.Value)' niet converteren (Index) ID $($message.Id)." }
                } # Fallback naar $message.hasAttachments is niet meer nodig/mogelijk

                # Creëer een object met de details van het huidige bericht
                $messageDetail = @{
                    MessageId        = $message.Id
                    Subject          = $message.Subject
                    ReceivedDateTime = $message.ReceivedDateTime
                    SenderName       = $emailSenderAddressInfo.Name # Naam van de afzender
                    SenderEmailAddress = $senderFullAddress # E-mailadres van de afzender
                    Size             = $currentMessageSize
                    HasAttachments   = $currentHasAttachments # Nieuw/bijgewerkt veld
                    ToRecipients     = $message.ToRecipients | ForEach-Object { $_.EmailAddress.Address }
                    Categories       = $message.Categories
                }

                if ($Script:SenderCache.ContainsKey($domainKey)) {
                    $Script:SenderCache[$domainKey].Count++
                    $Script:SenderCache[$domainKey].Messages.Add($messageDetail)
                } else {
                    $Script:SenderCache[$domainKey] = @{
                        Name     = $domainKey # Sla het domein op als 'Name' voor consistentie
                        Count    = 1
                        Messages = [System.Collections.Generic.List[PSObject]]::new()
                    }
                    $Script:SenderCache[$domainKey].Messages.Add($messageDetail)
                }
            }
        }
        Write-Progress -Activity "Mailbox Indexeren" -Completed

        $uniqueSenders = $Script:SenderCache.Keys.Count
        Write-Host "Indexeren voltooid. $uniqueSenders unieke afzenderdome(i)n(en) gevonden." # Aangepast voor domeinen

        Save-LocalCache # Sla de nieuw geïndexeerde cache op
    } catch {
        Write-Error "Fout tijdens het indexeren van de mailbox: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
        }
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
    # Read-Host "Druk op Enter om terug te keren naar het hoofdmenu" # Al eerder verwijderd
}

function Show-SenderOverview {
    param($UserId)

    # Definieer CGA-kleurenschema (consistent met Show-MainMenu)
    $cgaBgColor = [System.ConsoleColor]::Black
    $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green
    $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White
    $cgaWarningFgColor = [System.ConsoleColor]::Red
    # $cgaSpaceSelectedPrefixColor = [System.ConsoleColor]::Yellow # Niet gebruikt in dit menu

    # Functie-specifieke variabelen voor UI
    $selectedItemIndex = 0 # Index in de $sortedDomains array
    $topDisplayIndex = 0    # Index van het eerste domein dat getoond wordt in het venster
    $displayLines = 30      # Maximaal aantal domeinen tegelijk op het scherm

    # Hoofd lus voor dit menu
    $overviewLoopActive = $true
    while ($overviewLoopActive) {
        # Data laden en sorteren binnen de lus, zodat het ververst na acties
        if ($null -eq $Script:SenderCache -or $Script:SenderCache.Count -eq 0) {
            $Host.UI.RawUI.ForegroundColor = $cgaWarningFgColor
            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
            Clear-Host
            Write-Host "De mailbox is nog niet geïndexeerd of de index is leeg." -ForegroundColor $cgaWarningFgColor
            Write-Host "De automatische indexering bij het starten is mogelijk mislukt of er zijn geen gegevens gevonden. Controleer eventuele foutmeldingen bij het opstarten of probeer het script opnieuw." -ForegroundColor $cgaWarningFgColor
            $Host.UI.RawUI.ForegroundColor = $cgaInstructionFgColor
            Write-Host "Druk op Escape of Q om terug te keren."
            while($true){ $key = $Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown); if($key.VirtualKeyCode -eq 27 -or $key.Character.ToString().ToUpper() -eq 'Q'){ break } }
            return # Terug naar hoofdmenu
        }

        $domainList = @()
        foreach ($domainKey in $Script:SenderCache.Keys) {
            $domainList += [PSCustomObject]@{
                Domain = $domainKey
                Name   = $Script:SenderCache[$domainKey].Name # Dit is ook het domein
                Count  = $Script:SenderCache[$domainKey].Count
                Messages = $Script:SenderCache[$domainKey].Messages # Behoud de berichtenlijst voor acties
            }
        }
        $sortedDomains = $domainList | Sort-Object -Property @{Expression="Count"; Descending=$true}, Domain

        if ($sortedDomains.Count -eq 0) {
            $Host.UI.RawUI.ForegroundColor = $cgaFgColor
            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
            Clear-Host
            Write-Host "Geen domeinen (meer) gevonden in de cache."
            $Host.UI.RawUI.ForegroundColor = $cgaInstructionFgColor
            Write-Host "Druk op Escape of Q om terug te keren."
            while($true){ $key = $Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown); if($key.VirtualKeyCode -eq 27 -or $key.Character.ToString().ToUpper() -eq 'Q'){ break } }
            return # Terug naar hoofdmenu
        }

        # Zorg ervoor dat selectie en view binnen grenzen blijven na data herladen
        $selectedItemIndex = [Math]::Max(0, [Math]::Min($selectedItemIndex, $sortedDomains.Count - 1))
        $topDisplayIndex = [Math]::Max(0, [Math]::Min($topDisplayIndex, $sortedDomains.Count - $displayLines))
        if ($topDisplayIndex -lt 0) {$topDisplayIndex = 0}
        if ($selectedItemIndex -lt $topDisplayIndex) { $topDisplayIndex = $selectedItemIndex }
        if ($selectedItemIndex -ge ($topDisplayIndex + $displayLines)) { $topDisplayIndex = $selectedItemIndex - $displayLines + 1 }


        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host

        $title = "Overzicht van afzenderdomeinen (Scrollen: PgUp/PgDn/↑/↓, Enter: Open, V: Verplaats, Del: Verwijder, Esc/Q: Terug)"
        $headerLine = "{0,-5} {1,-7} {2,-50}" -f " ", "Aantal", "Domein" # Indicator kolom leeg voor nu
        $separator = "-" * ($Host.UI.RawUI.WindowSize.Width -1)

        Write-Host $title -ForegroundColor $cgaInstructionFgColor
        Write-Host $headerLine
        Write-Host $separator

        # Bepaal welke domeinen te tonen (voor paginering)
        $endDisplayIndex = [Math]::Min(($topDisplayIndex + $displayLines - 1), ($sortedDomains.Count - 1))

        for ($i = $topDisplayIndex; $i -le $endDisplayIndex; $i++) {
            $domainEntry = $sortedDomains[$i]

            $indicator = " " # Ruimte voor de ">" indicator
            $currentLineFgColor = $cgaFgColor
            $currentLineBgColor = $cgaBgColor

            if ($i -eq $selectedItemIndex) { # Huidig gehighlighte item
                $currentLineFgColor = $cgaSelectedFgColor
                $currentLineBgColor = $cgaSelectedBgColor
                $indicator = ">"
            }

            $itemText = "{0,-5} {1,-7} {2,-50}" -f $indicator, $domainEntry.Count, $domainEntry.Domain
            Write-Host $itemText -ForegroundColor $currentLineFgColor -BackgroundColor $currentLineBgColor
        }

        Write-Host $separator
        Write-Host ("Getoond: {0}-{1} van {2}" -f ($topDisplayIndex+1), ($endDisplayIndex+1), $sortedDomains.Count) -ForegroundColor $cgaInstructionFgColor

        # Wacht op toetsaanslag
        $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

        switch ($keyInfo.VirtualKeyCode) {
            38 { # UpArrow
                if ($selectedItemIndex -gt 0) { $selectedItemIndex-- }
                if ($selectedItemIndex -lt $topDisplayIndex) { $topDisplayIndex = $selectedItemIndex }
            }
            40 { # DownArrow
                if ($selectedItemIndex -lt ($sortedDomains.Count - 1)) { $selectedItemIndex++ }
                if ($selectedItemIndex -gt $endDisplayIndex) { $topDisplayIndex++ }
            }
            33 { # PageUp
                $selectedItemIndex = [Math]::Max(0, $selectedItemIndex - $displayLines)
                $topDisplayIndex = [Math]::Max(0, $topDisplayIndex - $displayLines)
                if ($selectedItemIndex -lt $topDisplayIndex) {$topDisplayIndex = $selectedItemIndex}
            }
            34 { # PageDown
                $selectedItemIndex = [Math]::Min(($sortedDomains.Count - 1), $selectedItemIndex + $displayLines)
                $topDisplayIndex = [Math]::Min(($sortedDomains.Count - $displayLines), $topDisplayIndex + $displayLines)
                if ($topDisplayIndex -lt 0) {$topDisplayIndex = 0}
                if ($selectedItemIndex -gt ($topDisplayIndex + $displayLines - 1)) {$topDisplayIndex = $selectedItemIndex - $displayLines + 1}
            }
            13 { # Enter - Open e-mails van dit domein
                if ($sortedDomains.Count -gt 0) {
                    $selectedDomainInfo = $sortedDomains[$selectedItemIndex]
                    Show-EmailsFromSelectedSender -UserId $UserId -SenderInfo $selectedDomainInfo
                }
            }
            86 { # V - Verplaats alle e-mails van dit domein
                if ($sortedDomains.Count -gt 0) {
                    $selectedDomainObject = $sortedDomains[$selectedItemIndex]
                    $messagesToActOn = $selectedDomainObject.Messages
                    if ($messagesToActOn -and $messagesToActOn.Count -gt 0) {
                        $Host.UI.RawUI.ForegroundColor = $cgaFgColor; $Host.UI.RawUI.BackgroundColor = $cgaBgColor
                        Perform-ActionOnAllSenderEmails -UserId $UserId -SenderDomain $selectedDomainObject.Domain -AllMessages $messagesToActOn -DirectAction "Move"
                    } else {
                        Write-Warning "Geen e-mails gevonden in de cache voor domein $($selectedDomainObject.Domain) om te verplaatsen."
                        Start-Sleep -Seconds 2
                    }
                }
            }
            46 { # Delete toets - Verwijder alle e-mails van dit domein
                 if ($sortedDomains.Count -gt 0) {
                    $selectedDomainObject = $sortedDomains[$selectedItemIndex]
                    $messagesToActOn = $selectedDomainObject.Messages
                    if ($messagesToActOn -and $messagesToActOn.Count -gt 0) {
                        $Host.UI.RawUI.ForegroundColor = $cgaFgColor; $Host.UI.RawUI.BackgroundColor = $cgaBgColor
                        Perform-ActionOnAllSenderEmails -UserId $UserId -SenderDomain $selectedDomainObject.Domain -AllMessages $messagesToActOn -DirectAction "Delete"
                    } else {
                        Write-Warning "Geen e-mails gevonden in de cache voor domein $($selectedDomainObject.Domain) om te verwijderen."
                        Start-Sleep -Seconds 2
                    }
                }
            }
            27 { $overviewLoopActive = $false } # Escape
            default {
                if ($keyInfo.Character.ToString().ToUpper() -eq 'Q') { $overviewLoopActive = $false }
            }
        }
    } # Einde while ($overviewLoopActive)
}

# Helper functie om HTML naar platte tekst te converteren
function Convert-HtmlToPlainText {
    param (
        [string]$HtmlContent
    )
    if ([string]::IsNullOrWhiteSpace($HtmlContent)) {
        return ""
    }
    # Verwijder script en style blokken eerst
    $plainText = $HtmlContent -replace '(?is)<script.*?</script>', '' -replace '(?is)<style.*?</style>', ''

    # Behandel specifieke tags voor newlines en basisopmaak
    # Headings (simulatie met extra newlines en prefix/suffix)
    $plainText = $plainText -replace '(?i)</h[1-6]>', "`r`n" # Newline na heading
    $plainText = $plainText -replace '(?i)<h1>(.*?)</h1>', "`r`n==== $1 ====`r`n"
    $plainText = $plainText -replace '(?i)<h2>(.*?)</h2>', "`r`n=== $1 ===`r`n"
    $plainText = $plainText -replace '(?i)<h3>(.*?)</h3>', "`r`n== $1 ==`r`n"
    $plainText = $plainText -replace '(?i)<h4>(.*?)</h4>', "`r`n= $1 =`r`n"
    # <p> tags
    $plainText = $plainText -replace '(?i)<p[^>]*>', "`r`n" # Start van <p> een newline
    $plainText = $plainText -replace '(?i)</p>', "`r`n"    # Einde van </p> ook een newline
    # <br> tags
    $plainText = $plainText -replace '(?i)<br\s*/?>', "`r`n"
    # <div> tags (behandel als paragraaf voor newlines)
    $plainText = $plainText -replace '(?i)<div[^>]*>', "`r`n"
    $plainText = $plainText -replace '(?i)</div>', "`r`n"
    # List items
    $plainText = $plainText -replace '(?i)<li[^>]*>', "`r`n  * " # Begin list item met newline en asterisk
    $plainText = $plainText -replace '(?i)</li>', "`r`n"       # Einde list item

    # Verwijder alle overige HTML tags (na de specifieke behandelingen)
    $plainText = $plainText -replace '<[^>]+>', ''

    # Decode HTML entities robuuster
    $plainText = [System.Net.WebUtility]::HtmlDecode($plainText)

    # Normaliseer witruimte
    # Vervang meerdere spaties door een enkele spatie (behalve newlines)
    $plainText = $plainText -replace '[ \t]{2,}', ' '
    # Trim spaties en tabs aan het begin/einde van elke regel
    $plainText = ($plainText.Split([string[]]@("`r`n", "`n"), [System.StringSplitOptions]::None) | ForEach-Object { $_.Trim() }) -join "`r`n"
    # Verwijder meerdere opeenvolgende lege regels, laat maximaal één lege regel toe
    $plainText = $plainText -replace "(\r?\n){3,}", "`r`n`r`n"

    return $plainText.Trim()
}

# Helper functie voor Ja/Nee bevestiging met pijltjesnavigatie
function Get-Confirmation {
    param (
        [string]$PromptMessage,
        [string]$WindowTitle = "Bevestiging"
    )

    # Sla huidige consolekleuren op (wordt hersteld door aanroepende functie of hoofdmenu)
    # $originalForegroundColor = $Host.UI.RawUI.ForegroundColor
    # $originalBackgroundColor = $Host.UI.RawUI.BackgroundColor

    # CGA-kleuren
    $cgaBgColor = [System.ConsoleColor]::Black
    $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green
    $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White
    $cgaWarningFgColor = [System.ConsoleColor]::Red # Voor de prompt message

    $options = @("Ja", "Nee")
    $selectedOptionIndex = 0 # Standaard "Ja"
    $confirmationLoopActive = $true

    while ($confirmationLoopActive) {
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        # Clear-Host is hier misschien te veel, we willen de context van de vraag behouden.
        # We tekenen de prompt en opties op de huidige cursorpositie of iets lager.
        # Voor nu, houden we het simpel en clearen we niet, de aanroeper moet de UI beheren.
        # Echter, voor een pop-up-achtig gevoel, zou je cursorpositie kunnen opslaan en herstellen.

        # Toon de prompt (mogelijk met een waarschuwingskleur)
        Write-Host $PromptMessage -ForegroundColor $cgaWarningFgColor

        # Toon de opties
        for ($i = 0; $i -lt $options.Count; $i++) {
            $optionText = $options[$i]
            if ($i -eq $selectedOptionIndex) {
                Write-Host "  > $($optionText)  " -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor -NoNewline
            } else {
                Write-Host "    $($optionText)  " -ForegroundColor $cgaFgColor -BackgroundColor $cgaBgColor -NoNewline
            }
        }
        Write-Host "" # Nieuwe regel na de opties

        $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

        switch ($keyInfo.VirtualKeyCode) {
            37 { # LeftArrow
                $selectedOptionIndex--
                if ($selectedOptionIndex -lt 0) { $selectedOptionIndex = $options.Count - 1 }
            }
            39 { # RightArrow
                $selectedOptionIndex++
                if ($selectedOptionIndex -ge $options.Count) { $selectedOptionIndex = 0 }
            }
            13 { # Enter
                $confirmationLoopActive = $false
                # Kleuren worden hersteld door de aanroepende functie
                return ($options[$selectedOptionIndex] -eq "Ja")
            }
            27 { # Escape
                $confirmationLoopActive = $false
                # Kleuren worden hersteld door de aanroepende functie
                return $false # Behandel Escape als "Nee" of annuleren
            }
        }
        # Wis de vorige opties (simpele aanpak: ga een regel omhoog en wis)
        # Dit is lastig zonder de exacte cursorpositie te beheren.
        # Voor nu, accepteren we dat de opties opnieuw worden getekend.
        # Een betere UI zou de cursorpositie opslaan en alleen de optieregels overschrijven.
        # Voor de eenvoud laten we dit nu achterwege. De aanroepende functie kan Clear-Host doen indien nodig.
    }
}


# Nieuwe helper functie om de cache bij te werken
function Update-SenderCache {
    param (
        [string]$DomainToUpdate, # Aangepast van SenderEmail naar DomainToUpdate
        [string]$MessageIdToRemove,
        [switch]$RemoveAllMessagesFromDomain # Aangepast van Sender naar Domain
    )

    # Als de update voor een speciale view is (zoals zoekresultaten), sla de cache update over.
    # Deze views beheren hun eigen lijsten of herladen data.
    if ($DomainToUpdate -eq "RECENT_EMAILS_VIEW" -or $DomainToUpdate -eq "SEARCH_RESULTS_VIEW") {
        Write-Verbose "Cache update overgeslagen voor speciale view: $DomainToUpdate"
        return
    }

    $normalizedDomainKey = $DomainToUpdate.ToLowerInvariant()

    if (-not $Script:SenderCache.ContainsKey($normalizedDomainKey)) {
        Write-Warning "Kan domein '$normalizedDomainKey' niet vinden in de cache voor update."
        return
    }

    if ($RemoveAllMessagesFromDomain) {
        Write-Host "Alle berichten van domein '$normalizedDomainKey' worden uit de cache verwijderd."
        $Script:SenderCache.Remove($normalizedDomainKey)
    } elseif ($MessageIdToRemove) {
        $messagesList = $Script:SenderCache[$normalizedDomainKey].Messages
        $messageToRemove = $messagesList | Where-Object { $_.MessageId -eq $MessageIdToRemove } | Select-Object -First 1

        if ($messageToRemove) {
            $messagesList.Remove($messageToRemove)
            $Script:SenderCache[$normalizedDomainKey].Count = $messagesList.Count
            Write-Host "Bericht met ID '$MessageIdToRemove' verwijderd uit cache voor domein '$normalizedDomainKey'. Nieuw aantal: $($messagesList.Count)."

            if ($messagesList.Count -eq 0) {
                Write-Host "Geen berichten meer voor domein '$normalizedDomainKey'. Domein wordt uit cache verwijderd."
                $Script:SenderCache.Remove($normalizedDomainKey)
            }
        } else {
            Write-Warning "Kon bericht met ID '$MessageIdToRemove' niet vinden in de cache voor domein '$normalizedDomainKey'."
        }
    }
    # Sla de cache op na elke update
    Save-LocalCache
}

# Nieuwe functie om e-mails van een geselecteerde afzender te tonen en acties te starten
function Show-EmailsFromSelectedSender {
    param (
        [string]$UserId,
        [PSCustomObject]$SenderInfo # Ontvangt nu een object met .Domain, .Name (is domein), .Count
    )

    $domainName = $SenderInfo.Domain
    $normalizedDomainKey = $domainName.ToLowerInvariant()

    # CGA Kleuren
    $cgaBgColor = [System.ConsoleColor]::Black
    $cgaFgColor = [System.ConsoleColor]::Green
    $cgaInstructionFgColor = [System.ConsoleColor]::White # Nodig voor callbacks

    if (-not $Script:SenderCache.ContainsKey($normalizedDomainKey) -or $Script:SenderCache[$normalizedDomainKey].Messages.Count -eq 0) {
        # Deze situatie zou idealiter al afgevangen moeten zijn in Show-SenderOverview
        # of de cache is gewijzigd tussen de selectie en deze aanroep.
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor # Herstel kleuren voor het geval ze anders waren
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host
        Write-Host "Geen e-mails (meer) in de cache voor domein '$domainName'." -ForegroundColor $cgaInstructionFgColor
        Write-Host "Druk op Escape of Q om terug te keren." -ForegroundColor $cgaInstructionFgColor
        $readKeyOptionsEmpty = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        while($true){ $key = $Host.UI.RawUI.ReadKey($readKeyOptionsEmpty); if($key.VirtualKeyCode -eq 27 -or $key.Character.ToString().ToUpper() -eq 'Q'){ break } }
        return # Verlaat Show-EmailsFromSelectedSender
    }

    $cachedDomainEntry = $Script:SenderCache[$normalizedDomainKey]
    $messagesFromCache = $cachedDomainEntry.Messages | Sort-Object ReceivedDateTime -Descending

    # Data voorbereiden voor Show-StandardizedEmailListView
    $messagesForView = @()
    foreach ($msgDetail in $messagesFromCache) {
        $messagesForView += [PSCustomObject]@{
            Id                 = $msgDetail.MessageId
            ReceivedDateTime   = $msgDetail.ReceivedDateTime
            Subject            = $msgDetail.Subject
            SenderName         = $msgDetail.SenderName
            SenderEmailAddress = $msgDetail.SenderEmailAddress
            Size               = $msgDetail.Size
            MessageForActions  = $msgDetail # Het cache (messageDetail) object
        }
    }

    # Callback functie om data te herladen (haalt opnieuw uit cache)
    $refreshCallbackForCache = {
        param($CurrentUserId, $CurrentNormalizedDomainKeyForCallback) # Context is de domain key
        Write-Host "E-maillijst voor domein '$CurrentNormalizedDomainKeyForCallback' herladen uit cache..." -ForegroundColor $cgaInstructionFgColor; Start-Sleep -Seconds 1
        $reloadedMessagesForView = @()
        if ($Script:SenderCache.ContainsKey($CurrentNormalizedDomainKeyForCallback)) {
            $reloadedCachedDomainEntry = $Script:SenderCache[$CurrentNormalizedDomainKeyForCallback]
            $reloadedMessagesFromCache = $reloadedCachedDomainEntry.Messages | Sort-Object ReceivedDateTime -Descending
            foreach ($rmsgDetail in $reloadedMessagesFromCache) {
                $reloadedMessagesForView += [PSCustomObject]@{
                    Id                 = $rmsgDetail.MessageId
                    ReceivedDateTime   = $rmsgDetail.ReceivedDateTime
                    Subject            = $rmsgDetail.Subject
                    SenderName         = $rmsgDetail.SenderName
                    SenderEmailAddress = $rmsgDetail.SenderEmailAddress
                    Size               = $rmsgDetail.Size
                    MessageForActions  = $rmsgDetail
                }
            }
        }
        # Als $reloadedMessagesForView leeg is na de refresh (bijv. alle mails verwijderd),
        # zal Show-StandardizedEmailListView dit correct afhandelen en terugkeren.
        return $reloadedMessagesForView
    }

    # Toon de e-maillijst met de generieke functie.
    # Show-StandardizedEmailListView handelt Esc/Q af en keert dan hier terug.
    # Na terugkeer uit Show-StandardizedEmailListView, zal deze functie Show-EmailsFromSelectedSender ook eindigen,
    # en de controle teruggeven aan Show-SenderOverview.
    Show-StandardizedEmailListView -UserId $UserId -Messages $messagesForView -ViewTitle "E-mails van domein: $($cachedDomainEntry.Name)" -AllowActions $true -DomainToUpdateCache $domainName -RefreshDataCallback $refreshCallbackForCache -RefreshDataCallbackContext $normalizedDomainKey
}

# NIEUWE GESTANDAARDISEERDE FUNCTIE VOOR E-MAILLIJSTEN
function Show-StandardizedEmailListView {
    param (
        [string]$UserId,
        [array]$Messages, # Array van PSCustomObjects met Id, ReceivedDateTime, Subject, SenderName, SenderEmailAddress, Size, MessageForActions
        [string]$ViewTitle,
        [bool]$AllowActions = $true, # Of acties zoals Verwijderen/Verplaatsen toegestaan zijn
        [string]$DomainToUpdateCache, # Voor cache updates na acties (kan een domeinnaam zijn of een speciale view identifier)
        [scriptblock]$RefreshDataCallback = $null, # Optionele callback om data te herladen
        $RefreshDataCallbackContext = $null # Optionele context voor de callback (bijv. $getMgUserMessageParams voor Search-Mail)
    )

    # CGA Kleuren
    $cgaBgColor = [System.ConsoleColor]::Black; $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green; $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White; $cgaWarningFgColor = [System.ConsoleColor]::Red
    $cgaSpaceSelectedPrefixColor = [System.ConsoleColor]::Yellow

    if (-not $Messages -or $Messages.Count -eq 0) {
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host
        Write-Host $ViewTitle -ForegroundColor $cgaInstructionFgColor
        Write-Host "Geen e-mails gevonden om weer te geven." -ForegroundColor $cgaInstructionFgColor
        Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
        $readKeyOptionsNoMsg = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        while ($Host.UI.RawUI.ReadKey($readKeyOptionsNoMsg).VirtualKeyCode -ne 27) {}
        return
    }

    $currentMessages = $Messages # Werk met een kopie die herladen kan worden
    $selectedEmailIndex = 0
    $topDisplayIndex = 0
    # Pas displayLines aan op basis van de vensterhoogte, met een minimum en aftrek voor headers/footers
    $displayLines = [Math]::Max(10, $Host.UI.RawUI.WindowSize.Height - 8) # -8 voor titel, instructies, headers, footer
    $spaceSelectedMessageIds = [System.Collections.Generic.HashSet[string]]::new()

    $emailListLoopActive = $true
    while ($emailListLoopActive) {
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host

        Write-Host "$ViewTitle (Scrollen: PgUp/PgDn/↑/↓, Spatie: Selecteer, Enter: Open/Acties)" -ForegroundColor $cgaInstructionFgColor
        if ($AllowActions) {
            Write-Host "Acties: V: Verplaats, Del: Verwijder | A: Selecteer Alles, N: Deselecteer Alles | Esc/Q: Terug" -ForegroundColor $cgaInstructionFgColor
        } else {
            Write-Host "A: Selecteer Alles, N: Deselecteer Alles | Esc/Q: Terug" -ForegroundColor $cgaInstructionFgColor
        }

        # Kolomvolgorde: Datum, Afzender Naam, Onderwerp, Afzender E-mail, Grootte
        # Pas breedtes aan voor leesbaarheid en om binnen $desiredWidth (150) te blijven
        # Totaal geschat: 5(#) + 20(Datum) + 25(Naam) + 45(Onderwerp) + 35(Email) + 15(Grootte) + (6*2=12 spaties) = 157. Lichte overschrijding, afkorten is belangrijk.
        $headerFormat = "{0,-5} {1,-20} {2,-25} {3,-45} {4,-35} {5,-15}"
        $lineFormat   = "{0} {1,-5} {2,-20} {3,-22} {4,-42} {5,-32} {6,-15}" # Prefix + ItemNr + Kolommen (iets kortere data strings)

        $separatorLine = "-" * ([Math]::Min(145, $Host.UI.RawUI.WindowSize.Width - 2)) # Max 145 tekens breed
        Write-Host $separatorLine
        Write-Host ($headerFormat -f "#", "Datum", "Afzender Naam", "Onderwerp", "Afzender E-mail", "Grootte (Bytes)")
        Write-Host $separatorLine

        $currentDisplayLines = [Math]::Min($displayLines, $currentMessages.Count)
        if ($currentMessages.Count -eq 0) { $currentDisplayLines = 0 }

        $endDisplayIndex = [Math]::Min(($topDisplayIndex + $currentDisplayLines - 1), ($currentMessages.Count - 1))
        if ($currentMessages.Count -eq 0) {
             Write-Host "Geen berichten (meer) om weer te geven." -ForegroundColor $cgaInstructionFgColor
        }

        for ($i = $topDisplayIndex; $i -le $endDisplayIndex; $i++) {
            if ($i -ge $currentMessages.Count) { break }
            $message = $currentMessages[$i]
            $itemNumber = $i + 1

            $receivedDisplay = if ($message.ReceivedDateTime) { Get-Date $message.ReceivedDateTime -Format "yyyy-MM-dd HH:mm" } else { "N/B" }

            $senderNameDisplay = if ($message.SenderName) { $message.SenderName } else { "N/B" }
            if ($senderNameDisplay.Length -gt 22) { $senderNameDisplay = $senderNameDisplay.Substring(0, 19) + "..." } # Inkorten

            $subjectDisplay = if ($message.Subject) { $message.Subject } else { "(Geen onderwerp)" }
            if ($subjectDisplay.Length -gt 42) { $subjectDisplay = $subjectDisplay.Substring(0, 39) + "..." } # Inkorten

            $senderEmailDisplay = if ($message.SenderEmailAddress) { $message.SenderEmailAddress } else { "N/B" }
            if ($senderEmailDisplay.Length -gt 32) { $senderEmailDisplay = $senderEmailDisplay.Substring(0, 29) + "..." } # Inkorten

            $sizeDisplay = if ($message.Size -ne $null) { $message.Size } else { "N/B" }

            $selectionPrefix = "   "
            $currentLineFgColor = $cgaFgColor
            $currentLineBgColor = $cgaBgColor

            if ($spaceSelectedMessageIds.Contains($message.Id)) {
                $selectionPrefix = "[*]"
            }

            if ($i -eq $selectedEmailIndex) {
                $currentLineFgColor = $cgaSelectedFgColor
                $currentLineBgColor = $cgaSelectedBgColor
                $selectionPrefix = if ($selectionPrefix -match "\[\*\]") { ">*]" } else { ">  " }
            }

            $lineText = $lineFormat -f $selectionPrefix, "$itemNumber.", $receivedDisplay, $senderNameDisplay, $subjectDisplay, $senderEmailDisplay, $sizeDisplay

            if (($selectionPrefix -match "\[\*\]" -or $selectionPrefix -match "^\>\*") -and $selectionPrefix.Length -ge 3) {
                Write-Host ($selectionPrefix.Substring(0,3)) -NoNewline -ForegroundColor $cgaSpaceSelectedPrefixColor -BackgroundColor $currentLineBgColor
                Write-Host ($lineText.Substring(3)) -ForegroundColor $currentLineFgColor -BackgroundColor $currentLineBgColor
            } else {
                Write-Host $lineText -ForegroundColor $currentLineFgColor -BackgroundColor $currentLineBgColor
            }
        }
        Write-Host $separatorLine
        $shownCountStart = if($currentMessages.Count -gt 0) { $topDisplayIndex + 1 } else { 0 }
        $shownCountEnd = if($currentMessages.Count -gt 0) { $endDisplayIndex + 1 } else { 0 }
        Write-Host ("Getoond: {0}-{1} van {2} | Geselecteerd (Spatie): {3}" -f $shownCountStart, $shownCountEnd, $currentMessages.Count, $spaceSelectedMessageIds.Count) -ForegroundColor $cgaInstructionFgColor

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
                    if ($selectedEmailIndex -gt ($topDisplayIndex + $currentDisplayLines - 1)) {$topDisplayIndex = $selectedEmailIndex - $currentDisplayLines + 1}
                }
            }
            32 { # Spacebar
                if ($currentMessages.Count -gt 0 -and $selectedEmailIndex -ge 0 -and $selectedEmailIndex -lt $currentMessages.Count) {
                    $currentMessageIdForSpace = $currentMessages[$selectedEmailIndex].Id
                    if ($spaceSelectedMessageIds.Contains($currentMessageIdForSpace)) {
                        $spaceSelectedMessageIds.Remove($currentMessageIdForSpace) | Out-Null
                    } else {
                        $spaceSelectedMessageIds.Add($currentMessageIdForSpace) | Out-Null
                    }
                }
            }
            13 { # Enter - Open email (of toon acties als het een cache object is)
                if ($currentMessages.Count -gt 0 -and $selectedEmailIndex -ge 0 -and $selectedEmailIndex -lt $currentMessages.Count) {
                    $Host.UI.RawUI.ForegroundColor = $cgaFgColor; $Host.UI.RawUI.BackgroundColor = $cgaBgColor
                    $selectedListViewItem = $currentMessages[$selectedEmailIndex]
                    $messageObjectToProcess = $selectedListViewItem.MessageForActions # Het Graph of Cache object
                    $knownGraphMessageId = $selectedListViewItem.Id # Dit is altijd de correcte Graph Message ID

                    # Roep altijd Show-EmailActionsMenu aan. Deze functie haalt de laatste versie van de mail op.
                    # Geef DomainToUpdateCache mee zodat Show-EmailActionsMenu de cache kan bijwerken.
                    Show-EmailActionsMenu -UserId $UserId -MessageId $knownGraphMessageId -DomainToUpdateCache $DomainToUpdateCache

                    if ($RefreshDataCallback) {
                        $currentMessages = Invoke-Command -ScriptBlock $RefreshDataCallback -ArgumentList $UserId, $RefreshDataCallbackContext
                        if (-not $currentMessages -or $currentMessages.Count -eq 0) { $emailListLoopActive = $false } else { $selectedEmailIndex = [Math]::Min($selectedEmailIndex, $currentMessages.Count -1); if ($selectedEmailIndex -lt 0) {$selectedEmailIndex = 0} }
                    }
                }
            }
            27 { $emailListLoopActive = $false } # Escape
            46 { # Delete toets
                if ($AllowActions) {
                    $messagesToActOnList = New-Object System.Collections.Generic.List[PSObject]
                    if ($spaceSelectedMessageIds.Count -gt 0) {
                        # Prioritize space-selected items
                        $currentMessages | Where-Object { $spaceSelectedMessageIds.Contains($_.Id) } | ForEach-Object { $messagesToActOnList.Add($_.MessageForActions) }
                    }
                    
                    # Fallback: If the list is still empty (either no space selection, or space selection yielded no matches)
                    # AND a single item is highlighted, then use the highlighted item.
                    if ($messagesToActOnList.Count -eq 0) {
                        if ($currentMessages.Count -gt 0 -and $selectedEmailIndex -ge 0 -and $selectedEmailIndex -lt $currentMessages.Count) {
                            $messagesToActOnList.Add($currentMessages[$selectedEmailIndex].MessageForActions)
                        }
                    }

                    if ($messagesToActOnList.Count -gt 0) {
                        $Host.UI.RawUI.ForegroundColor = $cgaFgColor; $Host.UI.RawUI.BackgroundColor = $cgaBgColor
                        Perform-ActionOnMultipleEmails -UserId $UserId -MessagesToProcess $messagesToActOnList -DomainToUpdateCache $DomainToUpdateCache -DirectAction "Delete"
                        $spaceSelectedMessageIds.Clear()
                        if ($RefreshDataCallback) {
                            $currentMessages = Invoke-Command -ScriptBlock $RefreshDataCallback -ArgumentList $UserId, $RefreshDataCallbackContext
                            if (-not $currentMessages -or $currentMessages.Count -eq 0) { $emailListLoopActive = $false } else { $selectedEmailIndex = [Math]::Min($selectedEmailIndex, $currentMessages.Count -1); if ($selectedEmailIndex -lt 0) {$selectedEmailIndex = 0} }
                        }
                    }
                }
            }
            86 { # V - Verplaatsen
                if ($AllowActions) {
                    $messagesToActOnList = New-Object System.Collections.Generic.List[PSObject]
                    if ($spaceSelectedMessageIds.Count -gt 0) {
                        # Prioritize space-selected items
                        $currentMessages | Where-Object { $spaceSelectedMessageIds.Contains($_.Id) } | ForEach-Object { $messagesToActOnList.Add($_.MessageForActions) }
                    }
                    
                    # Fallback: If the list is still empty (either no space selection, or space selection yielded no matches)
                    # AND a single item is highlighted, then use the highlighted item.
                    if ($messagesToActOnList.Count -eq 0) {
                        if ($currentMessages.Count -gt 0 -and $selectedEmailIndex -ge 0 -and $selectedEmailIndex -lt $currentMessages.Count) {
                            $messagesToActOnList.Add($currentMessages[$selectedEmailIndex].MessageForActions)
                        }
                    }

                    if ($messagesToActOnList.Count -gt 0) {
                        $Host.UI.RawUI.ForegroundColor = $cgaFgColor; $Host.UI.RawUI.BackgroundColor = $cgaBgColor
                        Perform-ActionOnMultipleEmails -UserId $UserId -MessagesToProcess $messagesToActOnList -DomainToUpdateCache $DomainToUpdateCache -DirectAction "Move"
                        $spaceSelectedMessageIds.Clear()
                        if ($RefreshDataCallback) {
                            $currentMessages = Invoke-Command -ScriptBlock $RefreshDataCallback -ArgumentList $UserId, $RefreshDataCallbackContext
                            if (-not $currentMessages -or $currentMessages.Count -eq 0) { $emailListLoopActive = $false } else { $selectedEmailIndex = [Math]::Min($selectedEmailIndex, $currentMessages.Count -1); if ($selectedEmailIndex -lt 0) {$selectedEmailIndex = 0} }
                        }
                    }
                }
            }
            default {
                $charPressed = $keyInfo.Character.ToString().ToUpper()

                if ($charPressed -eq 'A') { # A - Selecteer Alles
                    if ($currentMessages.Count -gt 0) {
                        foreach ($msg_sa in $currentMessages) {
                            $spaceSelectedMessageIds.Add($msg_sa.Id) | Out-Null
                        }
                    }
                } elseif ($charPressed -eq 'N') { # N - Deselecteer Alles (None)
                    $spaceSelectedMessageIds.Clear()
                } elseif ($charPressed -eq 'Q') { $emailListLoopActive = $false } # Q voor Quit
            }
        }

        if ($currentMessages -and $currentMessages.Count -gt 0) {
            $topDisplayIndex = [Math]::Max(0, [Math]::Min($topDisplayIndex, $currentMessages.Count - $currentDisplayLines))
            if ($topDisplayIndex -lt 0) {$topDisplayIndex = 0}
            $selectedEmailIndex = [Math]::Max(0, [Math]::Min($selectedEmailIndex, $currentMessages.Count - 1))
            if ($selectedEmailIndex -lt $topDisplayIndex) { $topDisplayIndex = $selectedEmailIndex }
            if ($selectedEmailIndex -ge ($topDisplayIndex + $currentDisplayLines) ) { $topDisplayIndex = $selectedEmailIndex - $currentDisplayLines + 1}
        } elseif (-not $currentMessages -or $currentMessages.Count -eq 0) {
            $emailListLoopActive = $false
        }
    }
}


# Nieuwe functie voor acties op meerdere (met spatie) geselecteerde e-mails
function Perform-ActionOnMultipleEmails {
    param (
        [string]$UserId,
        [System.Collections.Generic.List[PSObject]]$MessagesToProcess,
        [string]$DomainToUpdateCache,
        [ValidateSet("Delete", "Move")]
        [string]$DirectAction = $null # Nieuwe parameter voor directe actie
    )

    if ($MessagesToProcess.Count -eq 0) {
        Write-Warning "Geen berichten opgegeven om te verwerken."
        Start-Sleep -Seconds 2
        return
    }

    # CGA Kleuren
    $cgaBgColor = [System.ConsoleColor]::Black; $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green; $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White; $cgaWarningFgColor = [System.ConsoleColor]::Red

    $actionMenuItems = @(
        "1. Verwijder $($MessagesToProcess.Count) geselecteerde e-mail(s)",
        "2. Verplaats $($MessagesToProcess.Count) geselecteerde e-mail(s)",
        "3. Terug (Esc)"
    )
    $selectedActionItemIndex = 0
    $actionLoopActive = $true
    $actionToExecute = $null

    if ($DirectAction) {
        if ($DirectAction -eq "Delete") {
            $actionToExecute = $actionMenuItems[0] # "1. Verwijder..."
        } elseif ($DirectAction -eq "Move") {
            $actionToExecute = $actionMenuItems[1] # "2. Verplaats..."
        }
    } else { # Geen directe actie, toon menu
        while ($actionLoopActive) {
            $Host.UI.RawUI.ForegroundColor = $cgaFgColor
            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
            Clear-Host

            Write-Host "Acties voor $($MessagesToProcess.Count) geselecteerde e-mail(s):"
            Write-Host "-------------------------------------------"
            for ($i = 0; $i -lt $actionMenuItems.Count; $i++) {
                $itemText = $actionMenuItems[$i]
                if ($i -eq $selectedActionItemIndex) {
                    Write-Host "> $($itemText)" -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor
                } else {
                    Write-Host "  $($itemText)" -ForegroundColor $cgaFgColor
                }
            }
            Write-Host "Gebruik ↑/↓, Enter, Esc" -ForegroundColor $cgaInstructionFgColor

            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

            switch ($keyInfo.VirtualKeyCode) {
                38 { $selectedActionItemIndex = ($selectedActionItemIndex - 1 + $actionMenuItems.Count) % $actionMenuItems.Count }
                40 { $selectedActionItemIndex = ($selectedActionItemIndex + 1) % $actionMenuItems.Count }
                13 { $actionToExecute = $actionMenuItems[$selectedActionItemIndex]; $actionLoopActive = $false }
                27 { $actionToExecute = "3. Terug (Esc)"; $actionLoopActive = $false }
                default {
                    $char = $keyInfo.Character.ToString()
                    if ($char -eq '1') { $selectedActionItemIndex = 0; $actionToExecute = $actionMenuItems[0]; $actionLoopActive = $false }
                    if ($char -eq '2') { $selectedActionItemIndex = 1; $actionToExecute = $actionMenuItems[1]; $actionLoopActive = $false }
                    if ($char -eq '3') { $selectedActionItemIndex = 2; $actionToExecute = $actionMenuItems[2]; $actionLoopActive = $false }
                }
            }
        } # Einde while ($actionLoopActive)
    }


    if ($actionToExecute) {
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor; $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host

        if ($actionToExecute -like "1. Verwijder*") {
            if (Get-Confirmation -PromptMessage "Weet u zeker dat u deze $($MessagesToProcess.Count) e-mail(s) permanent wilt verwijderen?") {
                Write-Host "Starten met verwijderen van $($MessagesToProcess.Count) e-mail(s)..."
                $processedCount = 0; $errorCount = 0
                foreach ($message in $MessagesToProcess) {
                    $processedCount++
                    # Bepaal de daadwerkelijke Message ID
                    $effectiveMessageId = $null
                    if ($message.PSObject.Properties['Id'] -and -not [string]::IsNullOrWhiteSpace($message.Id)) {
                        $effectiveMessageId = $message.Id
                    } elseif ($message.PSObject.Properties['MessageId'] -and -not [string]::IsNullOrWhiteSpace($message.MessageId)) {
                        $effectiveMessageId = $message.MessageId
                    }

                    if ([string]::IsNullOrWhiteSpace($effectiveMessageId)) {
                        Write-Warning "Kon Message ID niet vinden voor bericht met onderwerp '$($message.Subject)'. Overslaan."
                        $errorCount++
                        continue
                    }

                    Write-Progress -Activity "Geselecteerde e-mails verwijderen" -Status "Verwijderen: $($message.Subject)" -PercentComplete (($processedCount / $MessagesToProcess.Count) * 100)
                    try {
                        Remove-MgUserMessage -UserId $UserId -MessageId $effectiveMessageId -ErrorAction Stop
                        Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $effectiveMessageId
                    } catch {
                        Write-Warning "Fout bij verwijderen e-mail ID $effectiveMessageId $($_.Exception.Message)"
                        $errorCount++
                    }
                }
                Write-Progress -Activity "Geselecteerde e-mails verwijderen" -Completed
                Write-Host "Verwijderen voltooid. $($MessagesToProcess.Count - $errorCount) e-mail(s) verwijderd."
                if ($errorCount -gt 0) { Write-Warning "$errorCount e-mail(s) konden niet worden verwijderd." }
                Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
                while($Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown).VirtualKeyCode -ne 27) {}
            } else {
                Write-Host "Verwijderen geannuleerd."
                Start-Sleep -Seconds 1 # Korte pauze om de melding te lezen
                # Keer direct terug, de lijstweergave zal verversen.
            }
        } elseif ($actionToExecute -like "2. Verplaats*") {
            $destinationFolderId = Get-MailFolderSelection -UserId $UserId
            if ($destinationFolderId) {
                $destinationFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $destinationFolderId -ErrorAction SilentlyContinue
                if (Get-Confirmation -PromptMessage "Weet u zeker dat u deze $($MessagesToProcess.Count) e-mail(s) wilt verplaatsen naar '$($destinationFolder.DisplayName)'?") {
                    Write-Host "Starten met verplaatsen van $($MessagesToProcess.Count) e-mail(s) naar '$($destinationFolder.DisplayName)'..."
                    $processedCount = 0; $errorCount = 0
                    foreach ($message in $MessagesToProcess) {
                        $processedCount++
                        # Bepaal de daadwerkelijke Message ID
                        $effectiveMessageId = $null
                        if ($message.PSObject.Properties['Id'] -and -not [string]::IsNullOrWhiteSpace($message.Id)) {
                            $effectiveMessageId = $message.Id
                        } elseif ($message.PSObject.Properties['MessageId'] -and -not [string]::IsNullOrWhiteSpace($message.MessageId)) {
                            $effectiveMessageId = $message.MessageId
                        }

                        if ([string]::IsNullOrWhiteSpace($effectiveMessageId)) {
                            Write-Warning "Kon Message ID niet vinden voor bericht met onderwerp '$($message.Subject)'. Overslaan."
                            $errorCount++
                            continue
                        }

                        Write-Progress -Activity "Geselecteerde e-mails verplaatsen" -Status "Verplaatsen: $($message.Subject)" -PercentComplete (($processedCount / $MessagesToProcess.Count) * 100)
                        try {
                            Move-MgUserMessage -UserId $UserId -MessageId $effectiveMessageId -DestinationId $destinationFolderId -ErrorAction Stop
                            Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $effectiveMessageId
                        } catch {
                            Write-Warning "Fout bij verplaatsen e-mail ID $effectiveMessageId $($_.Exception.Message)"
                            $errorCount++
                        }
                    }
                    Write-Progress -Activity "Geselecteerde e-mails verplaatsen" -Completed
                    Write-Host "Verplaatsen voltooid. $($MessagesToProcess.Count - $errorCount) e-mail(s) verplaatst."
                    if ($errorCount -gt 0) { Write-Warning "$errorCount e-mail(s) konden niet worden verplaatst." }
                    Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
                    while($Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown).VirtualKeyCode -ne 27) {}
                } else {
                    Write-Host "Verplaatsen geannuleerd."
                    Start-Sleep -Seconds 1 # Korte pauze om de melding te lezen
                }
            } else {
                Write-Host "Verplaatsen geannuleerd (geen doelmap geselecteerd)."
                Start-Sleep -Seconds 1 # Korte pauze om de melding te lezen
            }
        } elseif ($actionToExecute -like "3. Terug*") {
            # Do nothing, function will return
        }
    }
}

# Helper functie om de body van een e-mail te tonen
function Show-EmailBody {
    param (
        [string]$UserId,
        [PSCustomObject]$MessageObject, # Het volledige $messageDetail object uit de cache of direct van Graph
        [string]$KnownMessageId = $null # Optionele parameter voor een reeds bekende, betrouwbare Message ID
    )

    # CGA Kleuren (ervan uitgaande dat de aanroeper de kleuren beheert voor/na deze functie)
    $cgaBgColor = [System.ConsoleColor]::Black
    $cgaFgColor = [System.ConsoleColor]::Green
    $cgaInstructionFgColor = [System.ConsoleColor]::White

    $Host.UI.RawUI.ForegroundColor = $cgaFgColor
    $Host.UI.RawUI.BackgroundColor = $cgaBgColor
    Clear-Host

    # Bepaal de daadwerkelijke Message ID eigenschap
    $effectiveMessageId = $null
    if (-not [string]::IsNullOrWhiteSpace($KnownMessageId)) {
        $effectiveMessageId = $KnownMessageId
    } elseif ($MessageObject.PSObject.Properties['Id'] -and -not [string]::IsNullOrWhiteSpace($MessageObject.Id)) {
        $effectiveMessageId = $MessageObject.Id
    } elseif ($MessageObject.PSObject.Properties['MessageId'] -and -not [string]::IsNullOrWhiteSpace($MessageObject.MessageId)) {
        $effectiveMessageId = $MessageObject.MessageId
    }

    if ([string]::IsNullOrWhiteSpace($effectiveMessageId)) {
        Write-Error "Kan Message ID niet vinden (via KnownMessageId of in het opgegeven berichtobject)."
        Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
        $readKeyOptionsError = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        while ($Host.UI.RawUI.ReadKey($readKeyOptionsError).VirtualKeyCode -ne 27) {}
        return
    }

    Write-Host "Volledige body van e-mail:"
    Write-Host "Onderwerp    : $($MessageObject.Subject)"
    Write-Host "Ontvangen op : $(Get-Date $MessageObject.ReceivedDateTime -Format "yyyy-MM-dd HH:mm:ss")"
    Write-Host "ID           : $effectiveMessageId" # Gebruik de gevonden ID
    Write-Host "----------------------------------------------------"

    # Haal de volledige body op als die nog niet in $MessageObject zit (bijv. als het uit de cache komt zonder body)
    $bodyContent = ""
    $contentType = ""

    if ($MessageObject.PSObject.Properties["Body"] -and $MessageObject.Body.PSObject.Properties["Content"]) {
        $bodyContent = $MessageObject.Body.Content
        $contentType = $MessageObject.Body.ContentType
    } elseif ($MessageObject.PSObject.Properties["body"] -and $MessageObject.body.content) { # Soms is het lowercase
        $bodyContent = $MessageObject.body.content
        $contentType = $MessageObject.body.contentType
    }

    if ([string]::IsNullOrWhiteSpace($bodyContent)) {
        # Probeer de volledige body op te halen als deze ontbreekt of als we alleen een preview hadden
        Write-Host "Ophalen van volledige body van server..."
        try {
            $fullMessage = Get-MgUserMessage -UserId $UserId -MessageId $effectiveMessageId -Property "body" -ErrorAction Stop # Gebruik de gevonden ID
            if ($fullMessage -and $fullMessage.Body) {
                $bodyContent = $fullMessage.Body.Content
                $contentType = $fullMessage.Body.ContentType
            } else {
                $bodyContent = $MessageObject.BodyPreview # Fallback naar preview als body ophalen mislukt
                $contentType = "text" # Aanname
                if ([string]::IsNullOrWhiteSpace($bodyContent)) {
                    $bodyContent = "(Kon volledige body of preview niet ophalen)"
                }
            }
        } catch {
            Write-Error "Fout bij ophalen volledige body: $($_.Exception.Message)"
            $bodyContent = "(Fout bij ophalen body)"
            $contentType = "text" # Aanname
        }
    }

    if ($contentType -eq "html") {
        Write-Host "Originele body is HTML. Converteren naar platte tekst..."
        $displayText = Convert-HtmlToPlainText -HtmlContent $bodyContent
    } else {
        $displayText = $bodyContent
    }

    Write-Host $displayText
    Write-Host "----------------------------------------------------"
    Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor

    while ($true) {
        $readKeyOptionsBody = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown # Andere variabele naam
        $keyInfoBody = $Host.UI.RawUI.ReadKey($readKeyOptionsBody) # Andere variabele naam
        if ($keyInfoBody.VirtualKeyCode -eq 27) { # Escape
            break
        }
    }
    # Kleuren worden hersteld door de aanroepende functie
}

# Nieuwe functie voor acties op een enkele geselecteerde e-mail
function Perform-ActionOnSingleEmail {
    param (
        [string]$UserId,
        [PSCustomObject]$MessageObject, # Het $messageDetail object uit de cache
        [string]$DomainToUpdateCache # Aangepast van SenderEmailToUpdateCache
    )
    Clear-Host

    # Bepaal de daadwerkelijke Message ID eigenschap
    $effectiveMessageId = $null
    if ($MessageObject.PSObject.Properties['Id'] -and -not [string]::IsNullOrWhiteSpace($MessageObject.Id)) {
        $effectiveMessageId = $MessageObject.Id
    } elseif ($MessageObject.PSObject.Properties['MessageId'] -and -not [string]::IsNullOrWhiteSpace($MessageObject.MessageId)) {
        $effectiveMessageId = $MessageObject.MessageId
    }

    if ([string]::IsNullOrWhiteSpace($effectiveMessageId)) {
        Write-Error "Kan Message ID niet vinden in het opgegeven berichtobject voor Perform-ActionOnSingleEmail."
        Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
        $readKeyOptionsError = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        while ($Host.UI.RawUI.ReadKey($readKeyOptionsError).VirtualKeyCode -ne 27) {}
        return
    }

    Write-Host "Geselecteerde e-mail:"
    Write-Host "Onderwerp : $($MessageObject.Subject)"
    Write-Host "Ontvangen: $($MessageObject.ReceivedDateTime)"
    Write-Host "ID        : $effectiveMessageId"
    Write-Host "-------------------------------------------"

    # CGA Kleuren
    $cgaBgColor = [System.ConsoleColor]::Black; $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green; $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White

    $actionMenuItems = @(
        "1. Verwijder deze e-mail",
        "2. Verplaats deze e-mail",
        "3. Terug (Esc)"
    )
    $selectedActionItemIndex = 0
    $actionLoopActive = $true

    # Pre-laad details die constant blijven binnen de lus
    $emailSubjectDisplay = $MessageObject.Subject
    $emailReceivedDisplay = Get-Date $MessageObject.ReceivedDateTime -Format "yyyy-MM-dd HH:mm:ss"
    $emailIdDisplay = $effectiveMessageId # Gebruik de hierboven bepaalde ID

    while ($actionLoopActive) {
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host # Wis het scherm aan het begin van elke lus-iteratie

        # Herteken de statische informatie
        Write-Host "Geselecteerde e-mail:"
        Write-Host "Onderwerp : $emailSubjectDisplay"
        Write-Host "Ontvangen: $emailReceivedDisplay"
        Write-Host "ID        : $emailIdDisplay"
        Write-Host "-------------------------------------------"

        # Herteken het menu
        Write-Host "Kies een actie:" -ForegroundColor $cgaInstructionFgColor
        for ($i = 0; $i -lt $actionMenuItems.Count; $i++) {
            $itemText = $actionMenuItems[$i]
            if ($i -eq $selectedActionItemIndex) {
                Write-Host "> $($itemText)" -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor
            } else {
                Write-Host "  $($itemText)" -ForegroundColor $cgaFgColor
            }
        }

        $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)
        $actionToExecute = $null

        switch ($keyInfo.VirtualKeyCode) {
            38 { # UpArrow
                $selectedActionItemIndex--
                if ($selectedActionItemIndex -lt 0) { $selectedActionItemIndex = $actionMenuItems.Count - 1 }
            }
            40 { # DownArrow
                $selectedActionItemIndex++
                if ($selectedActionItemIndex -ge $actionMenuItems.Count) { $selectedActionItemIndex = 0 }
            }
            13 { # Enter
                $actionToExecute = $actionMenuItems[$selectedActionItemIndex]
            }
            27 { # Escape
                $actionToExecute = "3. Terug (Esc)" # Behandel Escape als Terug
            }
            default {
                $charPressed = $keyInfo.Character.ToString()
                if ($charPressed -eq '1') { $selectedActionItemIndex = 0; $actionToExecute = $actionMenuItems[0] }
                if ($charPressed -eq '2') { $selectedActionItemIndex = 1; $actionToExecute = $actionMenuItems[1] }
                if ($charPressed -eq '3') { $selectedActionItemIndex = 2; $actionToExecute = $actionMenuItems[2] }
            }
        }
        # Wis de vorige menu-opties (simpele aanpak: overschrijf met lege regels of Clear-Host als het ok is)
        # Voor nu, de lus zal het opnieuw tekenen.

        if ($actionToExecute) {
            $actionLoopActive = $false # Verlaat de lus na een keuze, tenzij het een sub-menu is

            # Herstel kleuren voor sub-acties
            $Host.UI.RawUI.ForegroundColor = $cgaFgColor
            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
            # Clear-Host # Of wis alleen de menu-opties

            if ($actionToExecute -like "1. Verwijder*") {
                if (Get-Confirmation -PromptMessage "Weet u zeker dat u deze e-mail permanent wilt verwijderen?") {
                    try {
                        Write-Host "Verwijderen van e-mail ID $effectiveMessageId..."
                    Remove-MgUserMessage -UserId $UserId -MessageId $effectiveMessageId -ErrorAction Stop
                    Write-Host "E-mail succesvol verwijderd van server."
                    # Update cache
                    Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $effectiveMessageId
                } catch {
                        Write-Error "Fout bij het verwijderen van e-mail ID $effectiveMessageId $($_.Exception.Message)"
                    }
                } else { Write-Host "Verwijderen geannuleerd." }
                # Read-Host "Druk op Enter om door te gaan." # Verwijderd
            } elseif ($actionToExecute -like "2. Verplaats*") {
                $destinationFolderId = Get-MailFolderSelection -UserId $UserId # Deze moet ook interactief worden
                if ($destinationFolderId) {
                    $destinationFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $destinationFolderId -ErrorAction SilentlyContinue
                    if (Get-Confirmation -PromptMessage "Weet u zeker dat u deze e-mail wilt verplaatsen naar '$($destinationFolder.DisplayName)'?") {
                        try {
                            Write-Host "Verplaatsen van e-mail ID $effectiveMessageId naar '$($destinationFolder.DisplayName)'..."
                            Move-MgUserMessage -UserId $UserId -MessageId $effectiveMessageId -DestinationId $destinationFolderId -ErrorAction Stop
                            Write-Host "E-mail succesvol verplaatst."
                        # Update cache
                            Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $effectiveMessageId
                        } catch {
                            Write-Error "Fout bij het verplaatsen van e-mail ID $effectiveMessageId $($_.Exception.Message)"
                        }
                    } else { Write-Host "Verplaatsen geannuleerd." }
                } else { Write-Host "Verplaatsen geannuleerd (geen doelmap geselecteerd)." }
                # Read-Host "Druk op Enter om door te gaan." # Verwijderd
            } elseif ($actionToExecute -like "3. Terug*") {
                # Do nothing, loop will exit
            } else {
                 Write-Warning "Ongeldige actie: $actionToExecute" # Zou niet moeten gebeuren
                 # Read-Host "Druk op Enter om door te gaan." # Verwijderd
            }
        } # end if ($actionToExecute)
    } # end while ($actionLoopActive)
    # Read-Host "Druk op Enter om terug te keren." # Verwijderd
}

# Nieuwe functie voor acties op ALLE e-mails van een afzender (vanuit de cache)
function Perform-ActionOnAllSenderEmails {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [string]$UserId,
        [string]$SenderDomain,
        [System.Collections.Generic.List[PSObject]]$AllMessages,
        [ValidateSet("Delete", "Move")]
        [string]$DirectAction = $null # Nieuwe parameter voor directe actie
    )

    # CGA Kleuren
    $cgaBgColor = [System.ConsoleColor]::Black; $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green; $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White
    $cgaWarningFgColor = [System.ConsoleColor]::Red

    $Host.UI.RawUI.ForegroundColor = $cgaFgColor
    $Host.UI.RawUI.BackgroundColor = $cgaBgColor
    Clear-Host

    Write-Host "Beheer ALLE e-mails van domein: $SenderDomain"
    Write-Host "Aantal te verwerken e-mails: $($AllMessages.Count)"
    Write-Host "-------------------------------------------"

    $actionMenuItems = @(
        "1. Verwijder ALLE e-mails van dit domein",
        "2. Verplaats ALLE e-mails van dit domein",
        "3. Terug (Esc/Q)"
    )
    $selectedActionItemIndex = 0
    $actionLoopActive = $true
    $actionToExecute = $null
    $allProcessedSuccessfully = $true # Standaard aanname, wordt false bij fouten

    # Pre-laad details die constant blijven binnen de lus
    $domainHeader = "Beheer ALLE e-mails van domein: $SenderDomain"
    $countHeader = "Aantal te verwerken e-mails: $($AllMessages.Count)"

    if ($DirectAction) {
        if ($DirectAction -eq "Delete") {
            $actionToExecute = $actionMenuItems[0] # "1. Verwijder..."
        } elseif ($DirectAction -eq "Move") {
            $actionToExecute = $actionMenuItems[1] # "2. Verplaats..."
        }
    } else { # Geen directe actie, toon menu
        while ($actionLoopActive) {
            $Host.UI.RawUI.ForegroundColor = $cgaFgColor
            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
            Clear-Host # Wis het scherm aan het begin van elke lus-iteratie

            # Herteken de statische informatie
            Write-Host $domainHeader
            Write-Host $countHeader
            Write-Host "-------------------------------------------"

            # Herteken het menu
            Write-Host "Kies een actie:" -ForegroundColor $cgaInstructionFgColor
            for ($i = 0; $i -lt $actionMenuItems.Count; $i++) {
                $itemText = $actionMenuItems[$i]
                if ($i -eq $selectedActionItemIndex) {
                    Write-Host "> $($itemText)" -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor
                } else {
                    Write-Host "  $($itemText)" -ForegroundColor $cgaFgColor
                }
            }
            Write-Host "Gebruik ↑/↓, Enter, Esc/Q" -ForegroundColor $cgaInstructionFgColor

            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

            switch ($keyInfo.VirtualKeyCode) {
                38 { # UpArrow
                    $selectedActionItemIndex--
                    if ($selectedActionItemIndex -lt 0) { $selectedActionItemIndex = $actionMenuItems.Count - 1 }
                }
                40 { # DownArrow
                    $selectedActionItemIndex++
                    if ($selectedActionItemIndex -ge $actionMenuItems.Count) { $selectedActionItemIndex = 0 }
                }
                13 { # Enter
                    $actionToExecute = $actionMenuItems[$selectedActionItemIndex]
                    $actionLoopActive = $false # Verlaat de keuzelus
                }
                27 { # Escape
                    $actionToExecute = "3. Terug (Esc/Q)"
                    $actionLoopActive = $false
                }
                default {
                    $charPressed = $keyInfo.Character.ToString().ToUpper()
                    if ($charPressed -eq '1') { $selectedActionItemIndex = 0; $actionToExecute = $actionMenuItems[0]; $actionLoopActive = $false }
                    if ($charPressed -eq '2') { $selectedActionItemIndex = 1; $actionToExecute = $actionMenuItems[1]; $actionLoopActive = $false }
                    if ($charPressed -eq '3' -or $charPressed -eq 'Q') { $selectedActionItemIndex = 2; $actionToExecute = $actionMenuItems[2]; $actionLoopActive = $false }
                }
            }
        } # Einde while ($actionLoopActive)
    }

    # Voer de gekozen actie uit
    if ($actionToExecute) {
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor # Herstel kleuren voor de actie output
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host # Maak scherm schoon voor de actie output

        if ($actionToExecute -like "1. Verwijder*") {
            if (Get-Confirmation -PromptMessage "WAARSCHUWING: Weet u zeker dat u ALLE $($AllMessages.Count) e-mails van domein '$SenderDomain' permanent wilt verwijderen?") {
                Write-Host "Starten met verwijderen van $($AllMessages.Count) e-mails van domein '$SenderDomain'..."
                $processedCount = 0
                $errorCount = 0
                foreach ($message in $AllMessages) {
                    $processedCount++
                    Write-Progress -Activity "Alle e-mails verwijderen" -Status "Verwijderen: $($message.Subject)" -PercentComplete (($processedCount / $AllMessages.Count) * 100)
                    try {
                        Remove-MgUserMessage -UserId $UserId -MessageId $message.MessageId -ErrorAction Stop
                    } catch {
                        Write-Warning "Fout bij verwijderen e-mail ID $($message.MessageId): $($_.Exception.Message)"
                        $errorCount++
                        $allProcessedSuccessfully = $false # Minstens één fout
                    }
                }
                Write-Progress -Activity "Alle e-mails verwijderen" -Completed
                Write-Host "Verwijderen voltooid. $($AllMessages.Count - $errorCount) e-mail(s) verwijderd."
                if ($errorCount -gt 0) { Write-Warning "$errorCount e-mail(s) konden niet worden verwijderd." }

                if ($allProcessedSuccessfully) { # Alleen als alles goed ging, verwijder de hele domein entry
                    Update-SenderCache -DomainToUpdate $SenderDomain -RemoveAllMessagesFromDomain # Aangepast
                    return $true # Signaleer dat de domein entry mogelijk weg is
                } elseif (($AllMessages.Count - $errorCount) -gt 0) { # Als sommigen zijn verwijderd, maar niet allen
                    # De cache moet individueel geüpdatet worden voor de succesvol verwijderde items.
                    # Dit is complexer; voor nu, informeer de gebruiker om opnieuw te indexeren.
                    Write-Warning "Niet alle e-mails konden worden verwijderd. De cache voor dit domein is mogelijk niet volledig accuraat. Indexeer opnieuw voor een correct overzicht."
                }
            } else { Write-Host "Verwijderen geannuleerd." }
            # Wacht op Escape om terug te keren
            Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            while($Host.UI.RawUI.ReadKey($readKeyOptions).VirtualKeyCode -ne 27) {}

        } elseif ($actionToExecute -like "2. Verplaats*") {
            $destinationFolderId = Get-MailFolderSelection -UserId $UserId # Deze moet ook interactief worden
            if ($destinationFolderId) {
                $destinationFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $destinationFolderId -ErrorAction SilentlyContinue
                if (Get-Confirmation -PromptMessage "WAARSCHUWING: Weet u zeker dat u ALLE $($AllMessages.Count) e-mails van domein '$SenderDomain' wilt verplaatsen naar '$($destinationFolder.DisplayName)'?") {
                    Write-Host "Starten met verplaatsen van $($AllMessages.Count) e-mails van domein '$SenderDomain' naar '$($destinationFolder.DisplayName)'..."
                    $processedCount = 0
                    $errorCount = 0
                    foreach ($message in $AllMessages) {
                        $processedCount++
                        Write-Progress -Activity "Alle e-mails verplaatsen" -Status "Verplaatsen: $($message.Subject)" -PercentComplete (($processedCount / $AllMessages.Count) * 100)
                        try {
                            Move-MgUserMessage -UserId $UserId -MessageId $message.MessageId -DestinationId $destinationFolderId -ErrorAction Stop
                        } catch {
                            Write-Warning "Fout bij verplaatsen e-mail ID $($message.MessageId): $($_.Exception.Message)"
                            $errorCount++
                            $allProcessedSuccessfully = $false
                        }
                    }
                    Write-Progress -Activity "Alle e-mails verplaatsen" -Completed
                    Write-Host "Verplaatsen voltooid. $($AllMessages.Count - $errorCount) e-mail(s) verplaatst."
                    if ($errorCount -gt 0) { Write-Warning "$errorCount e-mail(s) konden niet worden verplaatst." }

                    if ($allProcessedSuccessfully) {
                        Update-SenderCache -DomainToUpdate $SenderDomain -RemoveAllMessagesFromDomain # Aangepast
                        return $true # Signaleer dat de domein entry mogelijk weg is
                    } elseif (($AllMessages.Count - $errorCount) -gt 0) {
                         Write-Warning "Niet alle e-mails konden worden verplaatst. De cache voor dit domein is mogelijk niet volledig accuraat. Indexeer opnieuw voor een correct overzicht."
                    }
                } else { Write-Host "Verplaatsen geannuleerd." }
            } else { Write-Host "Verplaatsen geannuleerd (geen doelmap geselecteerd)." }
            # Wacht op Escape om terug te keren
            Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            while($Host.UI.RawUI.ReadKey($readKeyOptions).VirtualKeyCode -ne 27) {}

        } elseif ($actionToExecute -like "3. Terug*") {
            return $false # Terug, geen bulk actie uitgevoerd die de sender entry zou verwijderen
        } else {
            # Dit zou niet moeten gebeuren als $actionToExecute correct is ingesteld
            Write-Warning "Ongeldige actie: $actionToExecute"
            Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            while($Host.UI.RawUI.ReadKey($readKeyOptions).VirtualKeyCode -ne 27) {}
        }
    } else { # Geen actie gekozen (bijv. Escape in het actiemenu)
        return $false
    }

    # Read-Host "Druk op Enter om terug te keren." # Verwijderd
    return (-not $allProcessedSuccessfully) # Als er fouten waren, is de sender entry mogelijk nog (deels) relevant
}


function Manage-EmailsBySender {
    param($UserId)

    # CGA Kleuren
    $cgaBgColor = [System.ConsoleColor]::Black; $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green; $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White; $cgaWarningFgColor = [System.ConsoleColor]::Red

    $Host.UI.RawUI.ForegroundColor = $cgaFgColor
    $Host.UI.RawUI.BackgroundColor = $cgaBgColor
    Clear-Host
    Write-Host "Beheer e-mails van specifiek domein voor $UserId"
    Write-Host "----------------------------------------------------"

    if ($null -eq $Script:SenderCache -or $Script:SenderCache.Count -eq 0) {
        Write-Warning "De mailbox is nog niet geïndexeerd of de index is leeg."
        Write-Warning "Kies optie 'R. Ververs Index' in het hoofdmenu om de index op te bouwen."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    $domainNameInput = Read-Host "Voer het domein in van de afzender (bijv. example.com)"
    if ([string]::IsNullOrWhiteSpace($domainNameInput)) {
        Write-Warning "Geen domein ingevoerd. Actie geannuleerd."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    $normalizedDomainKey = $domainNameInput.ToLowerInvariant()

    if (-not $Script:SenderCache.ContainsKey($normalizedDomainKey)) {
        Write-Warning "Domein '$domainNameInput' niet gevonden in de cache. Controleer het domein of indexeer de mailbox opnieuw."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    $cachedDomainEntry = $Script:SenderCache[$normalizedDomainKey]
    if ($null -eq $cachedDomainEntry -or $cachedDomainEntry.Messages.Count -eq 0) {
        Write-Warning "Geen e-mails gevonden in de cache voor domein '$domainNameInput'."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    $messagesFromCache = $cachedDomainEntry.Messages | Sort-Object ReceivedDateTime -Descending

    # Data voorbereiden voor Show-StandardizedEmailListView
    $messagesForView = @()
    foreach ($msgDetail in $messagesFromCache) {
        $messagesForView += [PSCustomObject]@{
            Id                 = $msgDetail.MessageId
            ReceivedDateTime   = $msgDetail.ReceivedDateTime
            Subject            = $msgDetail.Subject
            SenderName         = $msgDetail.SenderName
            SenderEmailAddress = $msgDetail.SenderEmailAddress
            Size               = $msgDetail.Size
            MessageForActions  = $msgDetail # Het cache (messageDetail) object
        }
    }

    # Callback functie om data te herladen (haalt opnieuw uit cache)
    $refreshCallbackForCache = {
        param($CurrentUserId, $CurrentNormalizedDomainKeyForCallback) # Context is de domain key
        Write-Host "E-maillijst voor domein '$CurrentNormalizedDomainKeyForCallback' herladen uit cache..." -ForegroundColor $cgaInstructionFgColor; Start-Sleep -Seconds 1
        $reloadedMessagesForView = @()
        if ($Script:SenderCache.ContainsKey($CurrentNormalizedDomainKeyForCallback)) {
            $reloadedCachedDomainEntry = $Script:SenderCache[$CurrentNormalizedDomainKeyForCallback]
            $reloadedMessagesFromCache = $reloadedCachedDomainEntry.Messages | Sort-Object ReceivedDateTime -Descending
            foreach ($rmsgDetail in $reloadedMessagesFromCache) {
                $reloadedMessagesForView += [PSCustomObject]@{
                    Id                 = $rmsgDetail.MessageId
                    ReceivedDateTime   = $rmsgDetail.ReceivedDateTime
                    Subject            = $rmsgDetail.Subject
                    SenderName         = $rmsgDetail.SenderName
                    SenderEmailAddress = $rmsgDetail.SenderEmailAddress
                    Size               = $rmsgDetail.Size
                    MessageForActions  = $rmsgDetail
                }
            }
        }
        return $reloadedMessagesForView
    }

    # Toon de e-maillijst met de generieke functie
    Show-StandardizedEmailListView -UserId $UserId -Messages $messagesForView -ViewTitle "E-mails van domein: $($cachedDomainEntry.Name)" -AllowActions $true -DomainToUpdateCache $normalizedDomainKey -RefreshDataCallback $refreshCallbackForCache -RefreshDataCallbackContext $normalizedDomainKey

    # Na Show-StandardizedEmailListView keert de gebruiker terug naar het hoofdmenu.
    # De Read-Host hieronder is niet meer nodig, de hoofdmenu-lus handelt dit af.
    # Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
}

# Delete-MailsFromSender en Move-MailsFromSender zijn verwijderd.
# De functionaliteit wordt nu afgehandeld door Show-StandardizedEmailListView
# in combinatie met Perform-ActionOnMultipleEmails en Perform-ActionOnAllSenderEmails.

function Get-MailFolderSelection {
    param (
        [string]$UserId
    )

    # CGA Kleuren
    $cgaBgColor = [System.ConsoleColor]::Black; $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green; $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White; $cgaWarningFgColor = [System.ConsoleColor]::Red

    $Host.UI.RawUI.ForegroundColor = $cgaFgColor
    $Host.UI.RawUI.BackgroundColor = $cgaBgColor
    Clear-Host

    Write-Host "Selecteer een doelmap:"
    Write-Host "-----------------------"
    try {
        $allMailFolders = Get-MgUserMailFolder -UserId $UserId -All -ErrorAction Stop | Sort-Object DisplayName

        if ($null -eq $allMailFolders -or $allMailFolders.Count -eq 0) {
            Write-Warning "Geen mailmappen gevonden voor gebruiker $UserId."
            Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            while($Host.UI.RawUI.ReadKey($readKeyOptions).VirtualKeyCode -ne 27) {}
            return $null
        }

        # Bouw een lijst van mappen met hun volledige pad voor weergave
        $displayFolders = @()
        foreach ($folder in $allMailFolders) {
            $pathParts = @($folder.DisplayName)
            $currentParentId = $folder.ParentFolderId
            while ($currentParentId) {
                $parentFolder = $allMailFolders | Where-Object {$_.Id -eq $currentParentId} | Select-Object -First 1
                if ($parentFolder) {
                    $pathParts.Insert(0, $parentFolder.DisplayName)
                    $currentParentId = $parentFolder.ParentFolderId
                } else {
                    $currentParentId = $null # Ouder niet gevonden in de lijst, stop
                }
            }
            $displayFolders += [PSCustomObject]@{
                Id = $folder.Id
                DisplayPath = $pathParts -join " / "
            }
        }
        # Sorteer op het volledige pad
        $sortedDisplayFolders = $displayFolders | Sort-Object DisplayPath

        $selectedFolderIndex = 0
        $folderLoopActive = $true

        while ($folderLoopActive) {
            $Host.UI.RawUI.ForegroundColor = $cgaFgColor
            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
            Clear-Host # Herteken het hele menu elke keer

            Write-Host "Selecteer een doelmap:" -ForegroundColor $cgaInstructionFgColor
            Write-Host "----------------------------------------------------------------------"
            Write-Host ("{0,-5} {1}" -f "#", "Map Pad")
            Write-Host "----------------------------------------------------------------------"

            for ($i = 0; $i -lt $sortedDisplayFolders.Count; $i++) {
                $folderEntry = $sortedDisplayFolders[$i]
                $itemNumber = $i + 1
                $lineText = "{0,-5} {1}" -f "$itemNumber.", $folderEntry.DisplayPath

                if ($i -eq $selectedFolderIndex) {
                    Write-Host $lineText -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor
                } else {
                    Write-Host $lineText
                }
            }
            Write-Host "----------------------------------------------------------------------"
            Write-Host "Gebruik ↑/↓, Enter om te selecteren, Esc/Q om te annuleren." -ForegroundColor $cgaInstructionFgColor

            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

            switch ($keyInfo.VirtualKeyCode) {
                38 { # UpArrow
                    $selectedFolderIndex--
                    if ($selectedFolderIndex -lt 0) { $selectedFolderIndex = $sortedDisplayFolders.Count - 1 }
                }
                40 { # DownArrow
                    $selectedFolderIndex++
                    if ($selectedFolderIndex -ge $sortedDisplayFolders.Count) { $selectedFolderIndex = 0 }
                }
                13 { # Enter
                    return $sortedDisplayFolders[$selectedFolderIndex].Id
                }
                27 { # Escape
                    return $null # Annuleren
                }
                default {
                    $charPressed = $keyInfo.Character.ToString().ToUpper()
                    if ($charPressed -eq 'Q') {
                        return $null # Annuleren
                    } elseif ($charPressed -match "^\d+$") {
                        $numChoice = [int]$charPressed
                        if ($numChoice -ge 1 -and $numChoice -le $sortedDisplayFolders.Count) {
                            return $sortedDisplayFolders[$numChoice - 1].Id
                        }
                    }
                }
            }
        } # Einde while ($folderLoopActive)

    } catch {
        Write-Error "Fout bij het ophalen van mailmappen: $($_.Exception.Message)"
        Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
        $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        while($Host.UI.RawUI.ReadKey($readKeyOptions).VirtualKeyCode -ne 27) {}
        return $null
    }
}

function Search-Mail {
    param(
        [string]$UserId,
        [switch]$IsTestRun # Nieuwe parameter
    )
    Clear-Host
    Write-Host "Zoek naar e-mails in mailbox: $UserId"
    Write-Host "---------------------------------------"

    $searchTerm = Read-Host "Voer zoekterm in (hoofdletterongevoelig, gebruik * voor wildcards. Zoekt in onderwerp, body, afzender)"
    if ([string]::IsNullOrWhiteSpace($searchTerm)) {
        Write-Warning "Geen zoekterm ingevoerd. Zoekactie geannuleerd."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    try {
        Write-Host "Zoeken naar e-mails met term: '$searchTerm'..."
        # De -Search parameter gebruikt de Microsoft Search KQL syntax.
        # Standaard zoekt het in meerdere velden zoals onderwerp, body, afzender.
        $baseMessageProperties = "id,subject,from,receivedDateTime,bodyPreview" # Verwijder 'hasAttachments' en 'size'
        $messageSizeMapiPropertyId = "Integer 0x0E08" # PR_MESSAGE_SIZE
        $messageHasAttachMapiPropertyId = "Boolean 0x0E1B" # PR_HASATTACH
        $expandExtendedProperties = "singleValueExtendedProperties(`$filter=id eq '$messageSizeMapiPropertyId' or id eq '$messageHasAttachMapiPropertyId')"
        $foundMessages = $null

        $getMgUserMessageParams = @{
            UserId         = $UserId
            Search         = $searchTerm
            Top            = 100 # Beperk het aantal resultaten
            Property       = $baseMessageProperties
            ExpandProperty = $expandExtendedProperties
            ErrorAction    = "Stop"
        }

        if ($IsTestRun.IsPresent) {
            $getMgUserMessageParams.OrderBy = "receivedDateTime desc"
            Write-Host "(Testmodus actief: resultaten gesorteerd op nieuwste eerst)"
        }

        Write-Host "Zoekresultaten ophalen (incl. MAPI size)..."
        $foundMessages = Get-MgUserMessage @getMgUserMessageParams
        Write-Host "Zoekresultaten succesvol opgehaald."

        if ($null -eq $foundMessages -or $foundMessages.Count -eq 0) {
            Write-Host "Geen e-mails gevonden die overeenkomen met de zoekterm '$searchTerm'."
        } else {
            # CGA Kleuren (worden ingesteld door Show-StandardizedEmailListView)
            $cgaInstructionFgColor = [System.ConsoleColor]::White # Voor Write-Host binnen de callback

            $messagesForView = @()
            foreach ($msg in $foundMessages) {
                $currentMessageSize = $null
                $mapiSizeProp = $msg.SingleValueExtendedProperties | Where-Object { $_.Id -eq $messageSizeMapiPropertyId } | Select-Object -First 1
                if ($mapiSizeProp -and $mapiSizeProp.Value) {
                    try { $currentMessageSize = [long]$mapiSizeProp.Value } catch { Write-Verbose "Kon MAPI size ' $($mapiSizeProp.Value)' niet converteren (Zoeken) ID $($msg.Id)." }
                } # Fallback niet meer nodig

                $currentHasAttachments = $false
                $mapiAttachProp = $msg.SingleValueExtendedProperties | Where-Object { $_.Id -eq $messageHasAttachMapiPropertyId } | Select-Object -First 1
                if ($mapiAttachProp -and $mapiAttachProp.Value -ne $null) {
                    try { $currentHasAttachments = [System.Convert]::ToBoolean($mapiAttachProp.Value) } catch { Write-Verbose "Kon MAPI hasAttach '$($mapiAttachProp.Value)' niet converteren (Zoeken) ID $($msg.Id)." }
                } # Fallback niet meer nodig

                $messagesForView += [PSCustomObject]@{
                    Id                 = $msg.Id
                    ReceivedDateTime   = $msg.ReceivedDateTime
                    Subject            = $msg.Subject
                    SenderName         = if ($msg.From -and $msg.From.EmailAddress) { $msg.From.EmailAddress.Name } else { "N/B" }
                    SenderEmailAddress = if ($msg.From -and $msg.From.EmailAddress) { $msg.From.EmailAddress.Address } else { "N/B" }
                    Size               = $currentMessageSize
                    HasAttachments     = $currentHasAttachments # Toegevoegd
                    MessageForActions  = $msg # Het originele Graph object
                }
            }

            # Callback functie om data te herladen
            $refreshCallback = {
                param($CurrentUserId, $CallbackContext) # $CallbackContext is een hashtable met Params
                $CurrentGetMgUserMessageParamsForSearch = $CallbackContext.Params # Bevat al ExpandProperty
                $localMessageSizeMapiPropertyId = "Integer 0x0E08"
                $localMessageHasAttachMapiPropertyId = "Boolean 0x0E1B"

                Write-Host "Zoekresultaten herladen..." -ForegroundColor $cgaInstructionFgColor; Start-Sleep -Seconds 1
                $reloadedMessages = Get-MgUserMessage @CurrentGetMgUserMessageParamsForSearch -ErrorAction SilentlyContinue
                $reloadedMessagesForView = @()
                if ($reloadedMessages) {
                    foreach ($rmsg in $reloadedMessages) {
                        $reloadedMessageSize = $null
                        $reloadedMapiSizeProp = $rmsg.SingleValueExtendedProperties | Where-Object { $_.Id -eq $localMessageSizeMapiPropertyId } | Select-Object -First 1
                        if ($reloadedMapiSizeProp -and $reloadedMapiSizeProp.Value) {
                            try { $reloadedMessageSize = [long]$reloadedMapiSizeProp.Value } catch {}
                        } # Fallback niet meer nodig

                        $reloadedHasAttachments = $false
                        $reloadedMapiAttachProp = $rmsg.SingleValueExtendedProperties | Where-Object { $_.Id -eq $localMessageHasAttachMapiPropertyId } | Select-Object -First 1
                        if ($reloadedMapiAttachProp -and $reloadedMapiAttachProp.Value -ne $null) {
                            try { $reloadedHasAttachments = [System.Convert]::ToBoolean($reloadedMapiAttachProp.Value) } catch {}
                        } # Fallback niet meer nodig

                        $reloadedMessagesForView += [PSCustomObject]@{
                            Id                 = $rmsg.Id
                            ReceivedDateTime   = $rmsg.ReceivedDateTime
                            Subject            = $rmsg.Subject
                            SenderName         = if ($rmsg.From -and $rmsg.From.EmailAddress) { $rmsg.From.EmailAddress.Name } else { "N/B" }
                            SenderEmailAddress = if ($rmsg.From -and $rmsg.From.EmailAddress) { $rmsg.From.EmailAddress.Address } else { "N/B" }
                            Size               = $reloadedMessageSize
                            HasAttachments     = $reloadedHasAttachments # Toegevoegd
                            MessageForActions  = $rmsg
                        }
                    }
                }
                return $reloadedMessagesForView
            }

            $callbackContext = @{
                Params = $getMgUserMessageParams # Bevat nu Property en ExpandProperty
                # SizeUsed is niet meer nodig in de context
            }

            Show-StandardizedEmailListView -UserId $UserId -Messages $messagesForView -ViewTitle "Zoekresultaten voor '$searchTerm'" -AllowActions $true -DomainToUpdateCache "SEARCH_RESULTS_VIEW" -RefreshDataCallback $refreshCallback -RefreshDataCallbackContext $callbackContext
        }
    } catch {
        Write-Error "Fout tijdens het zoeken naar e-mails: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
    # Read-Host "Druk op Enter om terug te keren naar het hoofdmenu" # Verwijderd
}

function Show-RecentEmails {
    param($UserId)

    # CGA Kleuren
    $cgaBgColor = [System.ConsoleColor]::Black; $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green; $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White

    Clear-Host
    Write-Host "Ophalen van de laatste 100 e-mails voor $UserId..."

    try {
        $baseMessageProperties = "id,subject,from,receivedDateTime,bodyPreview" # Verwijder 'hasAttachments' en 'size'
        $messageSizeMapiPropertyId = "Integer 0x0E08" # PR_MESSAGE_SIZE
        $messageHasAttachMapiPropertyId = "Boolean 0x0E1B" # PR_HASATTACH
        $expandExtendedProperties = "singleValueExtendedProperties(`$filter=id eq '$messageSizeMapiPropertyId' or id eq '$messageHasAttachMapiPropertyId')"
        $recentMessages = $null

        $getMgUserMessageParams = @{
            UserId         = $UserId
            Top            = 100
            OrderBy        = "receivedDateTime desc"
            Property       = $baseMessageProperties
            ExpandProperty = $expandExtendedProperties
            ErrorAction    = "Stop"
        }

        Write-Host "Recente e-mails ophalen (incl. MAPI size)..."
        $recentMessages = Get-MgUserMessage @getMgUserMessageParams
        Write-Host "Recente e-mails succesvol opgehaald."

        if ($null -eq $recentMessages -or $recentMessages.Count -eq 0) {
            Write-Host "Geen recente e-mails gevonden."
        } else {
            $messagesForView = @()
            foreach ($msg in $recentMessages) {
                $currentMessageSize = $null
                $mapiSizeProp = $msg.SingleValueExtendedProperties | Where-Object { $_.Id -eq $messageSizeMapiPropertyId } | Select-Object -First 1
                if ($mapiSizeProp -and $mapiSizeProp.Value) {
                    try { $currentMessageSize = [long]$mapiSizeProp.Value } catch { Write-Verbose "Kon MAPI size ' $($mapiSizeProp.Value)' niet converteren (Recente) ID $($msg.Id)." }
                } # Fallback niet meer nodig

                $currentHasAttachments = $false
                $mapiAttachProp = $msg.SingleValueExtendedProperties | Where-Object { $_.Id -eq $messageHasAttachMapiPropertyId } | Select-Object -First 1
                if ($mapiAttachProp -and $mapiAttachProp.Value -ne $null) {
                    try { $currentHasAttachments = [System.Convert]::ToBoolean($mapiAttachProp.Value) } catch { Write-Verbose "Kon MAPI hasAttach '$($mapiAttachProp.Value)' niet converteren (Recente) ID $($msg.Id)." }
                } # Fallback niet meer nodig

                $messagesForView += [PSCustomObject]@{
                    Id                 = $msg.Id
                    ReceivedDateTime   = $msg.ReceivedDateTime
                    Subject            = $msg.Subject
                    SenderName         = if ($msg.From -and $msg.From.EmailAddress) { $msg.From.EmailAddress.Name } else { "N/B" }
                    SenderEmailAddress = if ($msg.From -and $msg.From.EmailAddress) { $msg.From.EmailAddress.Address } else { "N/B" }
                    Size               = $currentMessageSize
                    HasAttachments     = $currentHasAttachments # Toegevoegd
                    MessageForActions  = $msg # Het originele Graph object
                }
            }

            # Callback functie om data te herladen
            $refreshCallback = {
                param($CurrentUserId, $CallbackContext) # $CallbackContext is een hashtable met Params
                $CurrentGetMgUserMessageParamsForRecent = $CallbackContext.Params # Bevat al ExpandProperty
                $localMessageSizeMapiPropertyId = "Integer 0x0E08"
                $localMessageHasAttachMapiPropertyId = "Boolean 0x0E1B"

                Write-Host "Recente e-mails herladen..." -ForegroundColor $cgaInstructionFgColor; Start-Sleep -Seconds 1
                $reloadedMessages = Get-MgUserMessage @CurrentGetMgUserMessageParamsForRecent -ErrorAction SilentlyContinue
                $reloadedMessagesForView = @()
                if ($reloadedMessages) {
                    foreach ($rmsg in $reloadedMessages) {
                        $reloadedMessageSize = $null
                        $reloadedMapiSizeProp = $rmsg.SingleValueExtendedProperties | Where-Object { $_.Id -eq $localMessageSizeMapiPropertyId } | Select-Object -First 1
                        if ($reloadedMapiSizeProp -and $reloadedMapiSizeProp.Value) {
                            try { $reloadedMessageSize = [long]$reloadedMapiSizeProp.Value } catch {}
                        } # Fallback niet meer nodig

                        $reloadedHasAttachments = $false
                        $reloadedMapiAttachProp = $rmsg.SingleValueExtendedProperties | Where-Object { $_.Id -eq $localMessageHasAttachMapiPropertyId } | Select-Object -First 1
                        if ($reloadedMapiAttachProp -and $reloadedMapiAttachProp.Value -ne $null) {
                            try { $reloadedHasAttachments = [System.Convert]::ToBoolean($reloadedMapiAttachProp.Value) } catch {}
                        } # Fallback niet meer nodig

                        $reloadedMessagesForView += [PSCustomObject]@{
                            Id                 = $rmsg.Id
                            ReceivedDateTime   = $rmsg.ReceivedDateTime
                            Subject            = $rmsg.Subject
                            SenderName         = if ($rmsg.From -and $rmsg.From.EmailAddress) { $rmsg.From.EmailAddress.Name } else { "N/B" }
                            SenderEmailAddress = if ($rmsg.From -and $rmsg.From.EmailAddress) { $rmsg.From.EmailAddress.Address } else { "N/B" }
                            Size               = $reloadedMessageSize
                            HasAttachments     = $reloadedHasAttachments # Toegevoegd
                            MessageForActions  = $rmsg
                        }
                    }
                }
                return $reloadedMessagesForView
            }
            $callbackContext = @{
                Params = $getMgUserMessageParams # Bevat Property en ExpandProperty
                # SizeUsed is niet meer nodig
            }

            Show-StandardizedEmailListView -UserId $UserId -Messages $messagesForView -ViewTitle "Laatste 100 e-mails" -AllowActions $true -DomainToUpdateCache "RECENT_EMAILS_VIEW" -RefreshDataCallback $refreshCallback -RefreshDataCallbackContext $callbackContext
        }
    } catch {
        Write-Error "Fout bij het ophalen van recente e-mails: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
    # Read-Host "Druk op Enter om terug te keren naar het hoofdmenu" # Verwijderd
}

function Show-EmailActionsMenu {
    param(
        [string]$UserId,
        [string]$MessageId,
        [string]$DomainToUpdateCache = $null # Nieuwe parameter voor cache updates
    )
    Clear-Host

    try {
        # Haal de e-mail op met de benodigde details
        $properties = "subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments"
        $message = Get-MgUserMessage -UserId $UserId -MessageId $MessageId -Property $properties -ErrorAction Stop

        if (-not $message) {
            Write-Warning "Kan e-mail met ID '$MessageId' niet vinden."
            Read-Host "Druk op Enter om terug te keren"
            return
        }

        # Toon e-mail details
        Write-Host "Details voor e-mail:"
        Write-Host "----------------------------------------------------"
        Write-Host ("Onderwerp    : {0}" -f ($message.Subject | Out-String).Trim())
        Write-Host ("Van          : {0}" -f ($message.From.EmailAddress.Address | Out-String).Trim())
        Write-Host ("Aan          : {0}" -f (($message.ToRecipients | ForEach-Object {$_.EmailAddress.Address}) -join ", "))
        if ($message.CcRecipients) {
            Write-Host ("CC           : {0}" -f (($message.CcRecipients | ForEach-Object {$_.EmailAddress.Address}) -join ", "))
        }
        if ($message.BccRecipients) { # Meestal niet zichtbaar, maar Graph kan het soms retourneren afhankelijk van context
            Write-Host ("BCC          : {0}" -f (($message.BccRecipients | ForEach-Object {$_.EmailAddress.Address}) -join ", "))
        }
        Write-Host ("Ontvangen op : {0}" -f (Get-Date $message.ReceivedDateTime -Format "yyyy-MM-dd HH:mm:ss"))
        Write-Host ("Bijlagen     : {0}" -f ($message.HasAttachments | Out-String).Trim())
        Write-Host ("Preview      : {0}" -f ($message.BodyPreview | Out-String).Trim())
        Write-Host "----------------------------------------------------"
        Write-Host "ID           : $MessageId"
        Write-Host "----------------------------------------------------"

        # CGA Kleuren
        $cgaBgColor = [System.ConsoleColor]::Black; $cgaFgColor = [System.ConsoleColor]::Green
        $cgaInstructionFgColor = [System.ConsoleColor]::White

        $actionLoopActive = $true
        while ($actionLoopActive) {
            # Herteken details en acties elke keer, voor het geval subfuncties het scherm hebben gewijzigd
            $Host.UI.RawUI.ForegroundColor = $cgaFgColor
            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
            Clear-Host

            Write-Host "Details voor e-mail:"
            Write-Host "----------------------------------------------------"
            Write-Host ("Onderwerp    : {0}" -f ($message.Subject | Out-String).Trim())
            Write-Host ("Van          : {0}" -f ($message.From.EmailAddress.Address | Out-String).Trim())
            Write-Host ("Aan          : {0}" -f (($message.ToRecipients | ForEach-Object {$_.EmailAddress.Address}) -join ", "))
            if ($message.CcRecipients) {
                Write-Host ("CC           : {0}" -f (($message.CcRecipients | ForEach-Object {$_.EmailAddress.Address}) -join ", "))
            }
            if ($message.BccRecipients) {
                Write-Host ("BCC          : {0}" -f (($message.BccRecipients | ForEach-Object {$_.EmailAddress.Address}) -join ", "))
            }
            Write-Host ("Ontvangen op : {0}" -f (Get-Date $message.ReceivedDateTime -Format "yyyy-MM-dd HH:mm:ss"))
            Write-Host ("Bijlagen     : {0}" -f ($message.HasAttachments | Out-String).Trim())
            Write-Host ("Preview      : {0}" -f ($message.BodyPreview | Out-String).Trim())
            Write-Host "----------------------------------------------------"
            Write-Host "ID           : $MessageId"
            Write-Host "----------------------------------------------------"
            Write-Host ""
            Write-Host "Acties: [Del] Verwijder | [V] Verplaats | [B] Bekijk Body | [D] Download Bijlagen | [Esc/Q] Terug" -ForegroundColor $cgaInstructionFgColor

            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

            switch ($keyInfo.VirtualKeyCode) {
                46 { # Delete
                    if (Get-Confirmation -PromptMessage "Weet u zeker dat u deze e-mail permanent wilt verwijderen?") {
                        try {
                            Remove-MgUserMessage -UserId $UserId -MessageId $MessageId -ErrorAction Stop
                            Write-Host "E-mail succesvol verwijderd."
                            if ($DomainToUpdateCache) { # Controleer of er een cache is om te updaten
                                Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $MessageId
                            }
                        } catch { Write-Error "Fout bij het verwijderen van de e-mail: $($_.Exception.Message)" }
                        $actionLoopActive = $false # Verlaat de lus en functie
                    } # Anders (geen bevestiging), blijft de lus actief en wordt het menu opnieuw getoond
                }
                86 { # V - Verplaatsen
                    $destinationFolderId = Get-MailFolderSelection -UserId $UserId
                    if ($destinationFolderId) {
                        $destinationFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $destinationFolderId -ErrorAction SilentlyContinue
                        if (Get-Confirmation -PromptMessage "Weet u zeker dat u deze e-mail wilt verplaatsen naar '$($destinationFolder.DisplayName)'?") {
                            try {
                                Move-MgUserMessage -UserId $UserId -MessageId $MessageId -DestinationId $destinationFolderId -ErrorAction Stop
                                Write-Host "E-mail succesvol verplaatst naar '$($destinationFolder.DisplayName)'."
                                if ($DomainToUpdateCache) { # Controleer of er een cache is om te updaten
                                    Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $MessageId
                                }
                            } catch { Write-Error "Fout bij het verplaatsen van de e-mail: $($_.Exception.Message)" }
                            $actionLoopActive = $false # Verlaat de lus en functie
                        }
                    } else { Write-Host "Verplaatsen geannuleerd (geen doelmap geselecteerd)." ; Start-Sleep -Seconds 1 }
                }
                27 { $actionLoopActive = $false } # Escape
                default {
                    $charPressed = $keyInfo.Character.ToString().ToUpper()
                    if ($charPressed -eq 'B') { # Bekijk Body
                        Show-EmailBody -UserId $UserId -MessageObject $message -KnownMessageId $MessageId
                        # Na terugkeer uit Show-EmailBody, wordt de lus voortgezet en het menu opnieuw getekend.
                    } elseif ($charPressed -eq 'D') { # Download Bijlagen
                        if ($message.HasAttachments) {
                            Download-MessageAttachments -UserId $UserId -MessageId $MessageId -FullMessageObject $message
                        } else {
                            Write-Host "Deze e-mail heeft geen bijlagen." ; Start-Sleep -Seconds 1
                        }
                        # Na terugkeer uit Download-MessageAttachments, wordt de lus voortgezet.
                    } elseif ($charPressed -eq 'Q') {
                        $actionLoopActive = $false
                    }
                }
            }
        } # Einde while ($actionLoopActive)

    } catch {
        Write-Error "Fout bij het ophalen of verwerken van e-mailacties: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
    # Read-Host "Druk op Enter om terug te keren naar het hoofdmenu (of vorige menu indien van toepassing)" # Verwijderd
    # De functie keert nu direct terug na een actie of als de gebruiker 'Terug' kiest.
    # De aanroepende functie (Show-StandardizedEmailListView) zal de UI verder afhandelen.
}

function Ensure-DownloadPath {
    param (
        [string]$Path
    )
    if (-not (Test-Path -Path $Path)) {
        Write-Host "Aanmaken downloadmap: $Path"
        try {
            New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
        } catch {
            Write-Error "Kon downloadmap '$Path' niet aanmaken: $($_.Exception.Message)"
            return $false
        }
    }
    return $true
}

function Download-MessageAttachments {
    param(
        [string]$UserId,
        [string]$MessageId,
        [PSCustomObject]$FullMessageObject # Nieuwe parameter voor naamgevingsconventie
    )
    Clear-Host
    Write-Host "Bijlagen voor e-mail ID: $MessageId"
    Write-Host "-------------------------------------"

    try {
        $attachments = Get-MgUserMessageAttachment -UserId $UserId -MessageId $MessageId -ErrorAction Stop
        if ($null -eq $attachments -or $attachments.Count -eq 0) {
            Write-Warning "Geen bijlagen gevonden voor deze e-mail (ook al gaf HasAttachments 'true' aan)."
            # Read-Host "Druk op Enter om terug te keren" # Verwijderd
            Start-Sleep -Seconds 1 # Korte pauze zodat de gebruiker de melding kan lezen
            return
        }

        # CGA Kleuren
        $cgaBgColor = [System.ConsoleColor]::Black; $cgaFgColor = [System.ConsoleColor]::Green
        $cgaSelectedBgColor = [System.ConsoleColor]::Green; $cgaSelectedFgColor = [System.ConsoleColor]::Black
        $cgaInstructionFgColor = [System.ConsoleColor]::White

        $attachmentList = @($attachments) # Converteer naar array voor indexering
        $selectedAttachmentIndex = 0
        $downloadChoiceMade = $false
        $attachmentsToDownload = New-Object System.Collections.Generic.List[object]

        while (-not $downloadChoiceMade) {
            $Host.UI.RawUI.ForegroundColor = $cgaFgColor
            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
            Clear-Host
            Write-Host "Bijlagen voor e-mail ID: $MessageId"
            Write-Host "Onderwerp: $($FullMessageObject.Subject)"
            Write-Host "-------------------------------------"
            Write-Host "Beschikbare bijlagen:"
            for ($idx = 0; $idx -lt $attachmentList.Count; $idx++) {
                $att = $attachmentList[$idx]
                $line = "{0}. {1} ({2} bytes, Type: {3})" -f ($idx + 1), $att.Name, $att.Size, $att.ContentType
                if ($idx -eq $selectedAttachmentIndex) {
                    Write-Host ("> " + $line) -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor
                } else {
                    Write-Host ("  " + $line)
                }
            }
            Write-Host "-------------------------------------"
            Write-Host "[Enter] Download Geselecteerde | [A] Download ALLE | [Esc/Q] Annuleren" -ForegroundColor $cgaInstructionFgColor

            $readKeyOptionsAtt = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfoAtt = $Host.UI.RawUI.ReadKey($readKeyOptionsAtt)

            switch ($keyInfoAtt.VirtualKeyCode) {
                38 { # Up
                    if ($selectedAttachmentIndex -gt 0) { $selectedAttachmentIndex-- }
                }
                40 { # Down
                    if ($selectedAttachmentIndex -lt ($attachmentList.Count - 1)) { $selectedAttachmentIndex++ }
                }
                13 { # Enter
                    $attachmentsToDownload.Add($attachmentList[$selectedAttachmentIndex])
                    $downloadChoiceMade = $true
                }
                27 { # Escape
                    Write-Host "Downloaden geannuleerd." ; Start-Sleep -Seconds 1
                    return
                }
                default {
                    $charAtt = $keyInfoAtt.Character.ToString().ToUpper()
                    if ($charAtt -eq 'A') {
                        $attachmentList | ForEach-Object { $attachmentsToDownload.Add($_) }
                        $downloadChoiceMade = $true
                    } elseif ($charAtt -eq 'Q') {
                        Write-Host "Downloaden geannuleerd." ; Start-Sleep -Seconds 1
                        return
                    }
                }
            }
        }

        if ($attachmentsToDownload.Count -eq 0) { # Zou niet moeten gebeuren als $downloadChoiceMade true is
            Write-Host "Geen bijlagen geselecteerd voor download." ; Start-Sleep -Seconds 1
            return
        }

        $defaultDownloadPath = Join-Path -Path $PSScriptRoot -ChildPath "_attachments"
        $downloadPath = Read-Host "Voer het pad in voor de downloads (standaard: $defaultDownloadPath)"
        if ([string]::IsNullOrWhiteSpace($downloadPath)) {
            $downloadPath = $defaultDownloadPath
        }

        if (-not (Ensure-DownloadPath -Path $downloadPath)) {
            # Read-Host "Druk op Enter om terug te keren" # Verwijderd
            Start-Sleep -Seconds 1
            return
        }

        foreach ($attachment in $attachmentsToDownload) {
            # Implementeer nieuwe naamgevingsconventie
            $emailReceivedDate = Get-Date $FullMessageObject.ReceivedDateTime -Format "yyyy-MM-dd"

            $emailSenderAddress = $FullMessageObject.From.EmailAddress.Address
            $emailSenderDomain = "unknowndomain"
            $emailSenderExtension = "unknownext"
            if ($emailSenderAddress -and $emailSenderAddress -match "@") {
                $domainPart = ($emailSenderAddress -split '@')[1]
                if ($domainPart -match "\.") {
                    $emailSenderDomain = ($domainPart -split "\.", 2)[0]
                    $emailSenderExtension = ($domainPart -split "\.", 2)[1]
                } else {
                    $emailSenderDomain = $domainPart
                    $emailSenderExtension = ""
                }
            }

            $emailSubject = $FullMessageObject.Subject

            # Definieer ongeldige tekens voor bestandsnamen (inclusief pad-specifieke tekens)
            $invalidPathChars = [System.IO.Path]::GetInvalidFileNameChars() + @(':', '/', '\', '?', '*', '"', '<', '>', '|')
            $regexInvalidPathChars = "[{0}]" -f ([regex]::Escape(-join $invalidPathChars))

            $safeSubject = "NoSubject"
            if (-not [string]::IsNullOrWhiteSpace($emailSubject)) {
                $safeSubject = ($emailSubject -replace $regexInvalidPathChars, '_').Substring(0, [Math]::Min($emailSubject.Length, 30))
            }

            $attachmentOriginalName = $attachment.Name
            $attachmentBaseNamePart = [System.IO.Path]::GetFileNameWithoutExtension($attachmentOriginalName) -replace $regexInvalidPathChars, '_'
            $attachmentExtensionPart = [System.IO.Path]::GetExtension($attachmentOriginalName) # Inclusief de punt, bijv. ".pdf"

            $baseFileNameForSaving = "{0}_{1}_{2}-{3}-{4}" -f $emailReceivedDate, $emailSenderDomain, $emailSenderExtension, $safeSubject, $attachmentBaseNamePart

            # Zorg dat de volledige bestandsnaam (basis + extensie) niet te lang wordt
            $maxBaseNameLength = 200 # Arbitraire limiet voor de basisnaam om totale lengte te beheren
            if ($baseFileNameForSaving.Length -gt $maxBaseNameLength) {
                $baseFileNameForSaving = $baseFileNameForSaving.Substring(0, $maxBaseNameLength)
            }

            $currentFileNameAttempt = $baseFileNameForSaving + $attachmentExtensionPart
            $filePath = Join-Path -Path $downloadPath -ChildPath $currentFileNameAttempt
            $counter = 1

            # Controleer of bestand al bestaat en voeg teller toe indien nodig
            while (Test-Path $filePath) {
                $newFileNameWithCounter = "{0}_{1}{2}" -f $baseFileNameForSaving, $counter, $attachmentExtensionPart
                $filePath = Join-Path -Path $downloadPath -ChildPath $newFileNameWithCounter
                $counter++
            }

            Write-Host "Downloaden van '$($attachment.Name)' naar '$filePath'..."
            try {
                # Gebruik Invoke-MgGraphRequest om de raw content van de bijlage te krijgen
                # Zorg ervoor dat $attachment.Id correct is. Soms is het @odata.id of iets dergelijks.
                # Get-MgUserMessageAttachment retourneert objecten waar .Id de attachment ID is.
                $attachmentValueUri = "/users/$UserId/messages/$MessageId/attachments/$($attachment.Id)/`$value"
                $attachmentContentBytes = Invoke-MgGraphRequest -Method GET -Uri $attachmentValueUri -ErrorAction Stop -OutputType Binary # Vraag binaire output

                if ($attachmentContentBytes -and $attachmentContentBytes.Length -gt 0) {
                    [System.IO.File]::WriteAllBytes($filePath, $attachmentContentBytes)
                    Write-Host "Bijlage '$($attachment.Name)' succesvol opgeslagen als '$filePath'."
                } else {
                    Write-Warning "Invoke-MgGraphRequest gaf geen (of lege) content terug voor bijlage '$($attachment.Name)'. Overslaan."
                }
            } catch {
                Write-Warning "Fout bij downloaden of opslaan van bijlage '$($attachment.Name)': $($_.Exception.Message)"
            }
        }

    } catch {
        Write-Error "Fout bij het ophalen of downloaden van bijlagen: $($_.Exception.Message)"
    }
    # Read-Host "Druk op Enter om terug te keren" # Verwijderd, de aanroeper (Show-EmailActionsMenu) handelt de UI af.
}

function Empty-DeletedItemsFolder {
    param($UserId)
    Clear-Host
    Write-Host "Legen van de map 'Verwijderde Items' voor $UserId"
    Write-Host "----------------------------------------------------"

    $confirmation = Read-Host "WAARSCHUWING: Weet u zeker dat u ALLE items in de map 'Verwijderde Items' permanent wilt verwijderen? Deze actie kan niet ongedaan worden gemaakt. (ja/nee)"
    if ($confirmation -ne 'ja') {
        Write-Host "Legen van 'Verwijderde Items' geannuleerd."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    try {
        # Haal de 'deleteditems' folder op. Dit is een well-known name.
        $deletedItemsFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId "deleteditems" -ErrorAction Stop

        if (-not $deletedItemsFolder) {
            Write-Warning "Kon de map 'Verwijderde Items' niet vinden."
            Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
            return
        }

        Write-Host "Bezig met het legen van de map '$($deletedItemsFolder.DisplayName)'..."

        # Gebruik Invoke-MgEmptyUserMailFolder om de map te legen
        Invoke-MgEmptyUserMailFolder -UserId $UserId -MailFolderId $deletedItemsFolder.Id -ErrorAction Stop

        Write-Host "De map '$($deletedItemsFolder.DisplayName)' is succesvol geleegd."
        Write-Warning "De lokale cache (indien gebruikt voor 'Verwijderde Items', wat momenteel niet het geval is) is mogelijk niet meer accuraat. Overweeg opnieuw te indexeren indien nodig."

    } catch {
        Write-Error "Fout tijdens het legen van de map 'Verwijderde Items': $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
    # Read-Host "Druk op Enter om terug te keren naar het hoofdmenu" # Verwijderd
}

# Write-CenteredLine is niet meer nodig voor de nieuwe menu-uitlijning.
# De logica wordt direct in Show-MainMenu afgehandeld.

function Show-MainMenu {
    param (
        [string]$UserEmail
    )
    # Sla huidige consolekleuren op
    $originalForegroundColor = $Host.UI.RawUI.ForegroundColor
    $originalBackgroundColor = $Host.UI.RawUI.BackgroundColor

    # Definieer CGA-kleurenschema
    $cgaBgColor = [System.ConsoleColor]::Black
    $cgaFgColor = [System.ConsoleColor]::Green         # Lichtgroen voor standaard tekst
    $cgaSelectedBgColor = [System.ConsoleColor]::Green # Achtergrond voor geselecteerd item
    $cgaSelectedFgColor = [System.ConsoleColor]::Black # Tekstkleur voor geselecteerd item
    $cgaInstructionFgColor = [System.ConsoleColor]::White # Voor instructietekst

    # Menu-items en bijbehorende actiecodes
    $menuItems = @(
        "1. Overzicht van verzenders (uit cache)",
        "2. Beheer mails van specifieke afzender",    # Werkt op cache
        "3. Zoek naar een mail (live)",              # Zoekt live, niet uit cache
        "4. Bekijk laatste 100 e-mails (live)",     # Nieuwe optie
        "5. Leeg 'Verwijderde Items' (live)",
        "R. Ververs Index vanaf Server (Forceer Refresh)",
        "Q. Afsluiten"
    )
    $actionCodes = "1", "2", "3", "4", "5", "R", "Q"

    $selectedItemIndex = 0
    $menuLoopActive = $true

    while ($menuLoopActive) {
        # Stel CGA-kleuren in voor het menu
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host

        # Menu content
        $title = "MailCleanBuddy - Hoofdmenu voor $UserEmail"
        $separator = "------------------------------------------"
        $instructionText = "Gebruik ↑/↓ om te selecteren, Enter om te kiezen, Esc om af te sluiten."

        # Bouw de volledige menu-inhoud voor breedteberekening
        $tempMenuContentForWidth = @($title) + @($separator) + $menuItems + @($separator) + @($instructionText)

        $menuWidth = 0
        foreach ($line in $tempMenuContentForWidth) {
            if ($line.Length -gt $menuWidth) {
                $menuWidth = $line.Length
            }
        }
        $frameWidth = $menuWidth + 4
        $consoleWidth = $Host.UI.RawUI.WindowSize.Width
        $leftPaddingSpaces = [Math]::Max(0, ($consoleWidth - $frameWidth) / 2)
        $leftPadding = " " * $leftPaddingSpaces
        $innerFramePadding = "  "

        # Verticale padding
        $topPaddingLines = 3
        1..$topPaddingLines | ForEach-Object { Write-Host "" } # Gebruikt huidige achtergrondkleur

        # Teken titel en bovenste separator
        Write-Host ($leftPadding + $innerFramePadding + $title.PadRight($menuWidth) + $innerFramePadding)
        Write-Host ($leftPadding + $innerFramePadding + $separator.PadRight($menuWidth) + $innerFramePadding)

        # Teken menu-items
        for ($i = 0; $i -lt $menuItems.Count; $i++) {
            $itemText = $menuItems[$i]
            $lineContent = $innerFramePadding + $itemText.PadRight($menuWidth) + $innerFramePadding

            if ($i -eq $selectedItemIndex) {
                Write-Host ($leftPadding + $lineContent) -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor
            } else {
                Write-Host ($leftPadding + $lineContent) -ForegroundColor $cgaFgColor -BackgroundColor $cgaBgColor
            }
        }

        # Teken onderste separator en instructies
        Write-Host ($leftPadding + $innerFramePadding + $separator.PadRight($menuWidth) + $innerFramePadding)
        Write-Host ($leftPadding + $innerFramePadding + $instructionText.PadRight($menuWidth) + $innerFramePadding) -ForegroundColor $cgaInstructionFgColor

        # Wacht op toetsaanslag
        $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)
        $choiceToProcess = $null

        # Verwerk toetsaanslag
        switch ($keyInfo.VirtualKeyCode) {
            38 { # UpArrow
                $selectedItemIndex--
                if ($selectedItemIndex -lt 0) { $selectedItemIndex = $menuItems.Count - 1 }
            }
            40 { # DownArrow
                $selectedItemIndex++
                if ($selectedItemIndex -ge $menuItems.Count) { $selectedItemIndex = 0 }
            }
            13 { # Enter
                $choiceToProcess = $actionCodes[$selectedItemIndex]
            }
            27 { # Escape
                $choiceToProcess = "Q" # Behandel Escape als Afsluiten in hoofdmenu
            }
            default {
                # Verwerk directe numerieke/letterkeuze
                $charPressed = $keyInfo.Character.ToString().ToUpper()
                if ($actionCodes -contains $charPressed) {
                    $choiceToProcess = $charPressed
                    # Update selectedItemIndex om overeen te komen met de directe keuze
                    $selectedItemIndex = [array]::IndexOf($actionCodes, $charPressed)
                }
            }
        }

        if ($choiceToProcess) {
            # Herstel originele kleuren voordat een subactie wordt uitgevoerd
            $Host.UI.RawUI.ForegroundColor = $originalForegroundColor
            $Host.UI.RawUI.BackgroundColor = $originalBackgroundColor
            Clear-Host # Maak scherm schoon met originele kleuren voor de subactie

            switch ($choiceToProcess) {
                "1" { Show-SenderOverview -UserId $UserEmail }
                "2" { Manage-EmailsBySender -UserId $UserEmail }
                "3" { Search-Mail -UserId $UserEmail -IsTestRun:$TestRun.IsPresent }
                "4" { Show-RecentEmails -UserId $UserEmail } # Nieuwe actie
                "5" { Empty-DeletedItemsFolder -UserId $UserEmail }
                "R" {
                    Write-Host "Volledige indexering vanaf server wordt gestart..."
                    Index-Mailbox -UserId $UserEmail # Deze functie slaat de cache zelf op
                    Write-Host "Indexering voltooid. Druk op een toets om het menu opnieuw te laden."
                    $readKeyOptionsRefresh = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
                    $Host.UI.RawUI.ReadKey($readKeyOptionsRefresh) | Out-Null # Wacht op toetsaanslag
                }
                "Q" {
                    Write-Host "Afsluiten..."
                    $menuLoopActive = $false # Stop de menulus
                    # Kleuren worden hersteld door de finally block van het script of net hieronder
                }
                default {
                    # Dit zou niet moeten gebeuren als $choiceToProcess correct is ingesteld
                    Write-Warning "Onbekende actie: $choiceToProcess"
                    Read-Host "Druk op Enter om door te gaan"
                }
            }

            if (-not $menuLoopActive) { # Als 'Q' is gekozen
                 # Herstel kleuren expliciet hier ook voor het geval de finally block niet direct volgt
                $Host.UI.RawUI.ForegroundColor = $originalForegroundColor
                $Host.UI.RawUI.BackgroundColor = $originalBackgroundColor
                return $false # Signaleer om de hoofdscriptlus te stoppen
            }
            # Na de subactie (als niet Q), worden CGA-kleuren opnieuw ingesteld aan het begin van de volgende while-lus iteratie.
        }
    } # Einde while ($menuLoopActive)

    # Mocht de lus op een andere manier eindigen (zou niet moeten), herstel kleuren.
    $Host.UI.RawUI.ForegroundColor = $originalForegroundColor
    $Host.UI.RawUI.BackgroundColor = $originalBackgroundColor
    return $true # Standaard, als de lus eindigt, ga door (hoewel Q $false retourneert)
}

try {
    # Define required Graph API scopes
    $RequiredScopes = @("Mail.Read", "User.Read", "Mail.ReadWrite") # Mail.ReadWrite for delete/move operations

    # Check, install if necessary, and import Microsoft.Graph.Authentication module
    try {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            Write-Host "Microsoft.Graph.Authentication module not found. Attempting to install..."
            Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -Confirm:$false -ErrorAction Stop
            Write-Host "Microsoft.Graph.Authentication module installed."
        }
        Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop # Added -Force
        Write-Host "Microsoft.Graph.Authentication module loaded successfully."
    }
    catch {
        throw "Kritiek: Kon de Microsoft.Graph.Authentication module niet installeren of importeren. Installeer deze handmatig met 'Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' en probeer het script opnieuw. Foutdetails: $($_.Exception.Message)"
    }

    # Check, install if necessary, and import Microsoft.Graph.Mail module
    try {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Mail)) {
            Write-Host "Microsoft.Graph.Mail module not found. Attempting to install..."
            Install-Module Microsoft.Graph.Mail -Scope CurrentUser -Force -Confirm:$false -ErrorAction Stop
            Write-Host "Microsoft.Graph.Mail module installed."
        }
        Import-Module Microsoft.Graph.Mail -Force -ErrorAction Stop # Added -Force
        Write-Host "Microsoft.Graph.Mail module loaded successfully."
    }
    catch {
        throw "Kritiek: Kon de Microsoft.Graph.Mail module niet installeren of importeren. Installeer deze handmatig met 'Install-Module Microsoft.Graph.Mail -Scope CurrentUser' en probeer het script opnieuw. Foutdetails: $($_.Exception.Message)"
    }

    # Connect to Microsoft Graph
    Write-Host "Attempting to connect to Microsoft Graph for mailbox: $MailboxEmail"
    try {
        # Check current connection and scopes
        $currentConnection = Get-MgContext -ErrorAction SilentlyContinue
        $hasRequiredScopes = $false
        if ($currentConnection) {
            $scopesMatch = $true
            foreach ($scope in $RequiredScopes) {
                if ($currentConnection.Scopes -notcontains $scope) {
                    $scopesMatch = $false
                    break
                }
            }
            if ($scopesMatch -and ($currentConnection.Scopes.Count -eq $RequiredScopes.Count)) {
                 $hasRequiredScopes = $true
            }
        }

        if (-not $currentConnection -or -not $hasRequiredScopes) {
            if ($currentConnection -and -not $hasRequiredScopes) {
                Write-Warning "Current Graph connection does not have all required scopes. Reconnecting."
                Disconnect-MgGraph -ErrorAction SilentlyContinue
            }
            Write-Host "Connecting to Microsoft Graph with scopes: $($RequiredScopes -join ', ')"
            Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop
        } else {
            Write-Host "Already connected to Microsoft Graph with required scopes."
        }
        Write-Host "Successfully connected to Microsoft Graph."

        # Verify that Graph cmdlets are available (optional, Connect-MgGraph success usually implies this)
        if (-not (Get-Command Get-MgUserMessage -ErrorAction SilentlyContinue)) {
            throw "Kritiek: Get-MgUserMessage cmdlet is niet beschikbaar na een succesvolle verbinding met Microsoft Graph. Controleer de Microsoft.Graph.Mail module."
        }
        Write-Host "Verbinding succesvol."

        # Initialiseer en laad/bouw de cache
        Get-CacheFilePath -MailboxEmail $MailboxEmail
        if (Load-LocalCache) {
            Write-Host "Lokale cache geladen. Volledige indexering wordt overgeslagen voor snellere start."
            Write-Host "Gebruik menu-optie 'R' om de index handmatig vanaf de server te verversen."
            # $Script:SenderCache is nu gevuld vanuit het bestand
        } else {
            Write-Host "Geen (valide) lokale cache gevonden. Starten met automatische indexering van de mailbox..."
            Index-Mailbox -UserId $MailboxEmail # Indexeert en slaat de cache op
            Write-Host "Automatische indexering voltooid (of poging daartoe)."
        }
        # Start-Sleep -Seconds 1 # Korte pauze
    }
    catch {
        throw "Kritiek: Fout tijdens het verbinden met Microsoft Graph of initialisatie van de cache: $($_.Exception.Message). Controleer de internetverbinding, de Microsoft Graph module installaties en de benodigde rechten/consent."
    }

    # Main application loop
    $keepRunning = $true
    while ($keepRunning) {
        $keepRunning = Show-MainMenu -UserEmail $MailboxEmail
    }

}
catch {
    Write-Error "Er is een fout opgetreden: $($_.Exception.Message)"
    if ($_.ScriptStackTrace) {
        Write-Error "StackTrace: $($_.ScriptStackTrace)"
    }
    if ($_.Exception.InnerException) {
        Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
    }
}
finally {
    # Disconnect from Microsoft Graph
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Write-Host "Disconnecting from Microsoft Graph..."
        Disconnect-MgGraph
    } else {
        Write-Host "Not connected to Microsoft Graph, or context is unavailable. No disconnection needed."
    }
}
