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
    .\OutlookBuddy.ps1 -MailboxEmail "user@example.com"
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
    [switch]$TestRun
)

# Script-level cache for sender information
$Script:SenderCache = $null 

# Placeholder functions for menu items
function Index-Mailbox {
    param($UserId)
    
    Write-Host "Starten met indexeren van mailbox voor $UserId..."
    if ($TestRun.IsPresent) {
        Write-Warning "** TESTMODUS ACTIEF: Alleen de laatste 100 e-mails worden geïndexeerd. **"
    }
    $Script:SenderCache = @{} # Reset of initialiseer de cache

    try {
        Write-Host "Ophalen van berichten..."
        $baseMessageProperties = "id", "subject", "sender", "receivedDateTime", "toRecipients", "categories"
        $sizeProperty = "Size" # De eigenschap die mogelijk problemen veroorzaakt
        
        $messages = $null
        $sizePropertySuccessfullyUsed = $true # Standaard aanname

        try {
            # Poging 1: Berichten ophalen inclusief de 'Size' eigenschap
            $currentMessageProperties = $baseMessageProperties + $sizeProperty
            Write-Host "Poging 1: Berichten ophalen inclusief '$sizeProperty' eigenschap..."
            if ($TestRun.IsPresent) {
                $messages = Get-MgUserMessage -UserId $UserId -Top 100 -Property $currentMessageProperties -OrderBy "receivedDateTime desc" -ErrorAction Stop
                Write-Host "(Testmodus: max 100 berichten opgehaald met '$sizeProperty')"
            } else {
                Write-Host "(Volledige modus met '$sizeProperty': dit kan even duren voor grote mailboxen)..."
                $messages = Get-MgUserMessage -UserId $UserId -All -Property $currentMessageProperties -ErrorAction Stop
            }
            Write-Host "Berichten succesvol opgehaald inclusief '$sizeProperty'."
        }
        catch {
            # Controleer of de specifieke fout met betrekking tot 'size' is opgetreden
            $errorMessage = $_.Exception.Message
            if ($_.Exception.InnerException) {
                $errorMessage = $_.Exception.InnerException.Message
            }

            if ($errorMessage -like "*Could not find a property named 'size' on type 'Microsoft.OutlookServices.Message'*") {
                Write-Warning "Fout bij ophalen berichten met eigenschap '$sizeProperty': $errorMessage"
                Write-Host "Poging 2: Berichten ophalen ZONDER '$sizeProperty' eigenschap..."
                $sizePropertySuccessfullyUsed = $false
                
                # Poging 2: Berichten ophalen ZONDER de 'Size' eigenschap
                if ($TestRun.IsPresent) {
                    $messages = Get-MgUserMessage -UserId $UserId -Top 100 -Property $baseMessageProperties -OrderBy "receivedDateTime desc" -ErrorAction Stop
                    Write-Host "(Testmodus: max 100 berichten opgehaald zonder '$sizeProperty')"
                } else {
                    Write-Host "(Volledige modus zonder '$sizeProperty': dit kan even duren voor grote mailboxen)..."
                    $messages = Get-MgUserMessage -UserId $UserId -All -Property $baseMessageProperties -ErrorAction Stop
                }
                Write-Host "Berichten succesvol opgehaald zonder '$sizeProperty'. Grootte-informatie zal ontbreken of leeg zijn."
            }
            else {
                # Een andere, onverwachte fout is opgetreden, gooi deze opnieuw om door de buitenste catch te worden afgehandeld
                throw $_ 
            }
        }
        
        if ($null -eq $messages -or $messages.Count -eq 0) {
            Write-Warning "Geen berichten gevonden in de mailbox."
            Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
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

            $sender = $message.Sender.EmailAddress
            if ($sender -and $sender.Address) {
                # Groepeer op domein
                $senderFullAddress = $sender.Address
                $domain = ($senderFullAddress -split '@')[1]
                if ([string]::IsNullOrWhiteSpace($domain)) {
                    $domain = "onbekend_domein" # Fallback voor ongeldige e-mailadressen
                }
                $domainKey = $domain.ToLowerInvariant()
                # De 'naam' voor de cache entry wordt nu het domein zelf.
                # De oorspronkelijke $senderName ($sender.Name) wordt niet meer direct gebruikt voor de groepering.

                # Bepaal de grootte van het bericht, afhankelijk van of de 'Size' eigenschap succesvol kon worden opgevraagd
                $currentMessageSize = $null
                if ($sizePropertySuccessfullyUsed) {
                    # Als 'Size' werd opgevraagd, probeer de waarde ervan te lezen.
                    # Controleer of de eigenschap 'Size' bestaat op het $message object om fouten te voorkomen.
                    if ($message.PSObject.Properties['Size']) {
                        $currentMessageSize = $message.Size
                    }
                    # Als de eigenschap niet bestaat op dit specifieke bericht (ondanks dat het was opgevraagd), blijft $currentMessageSize $null.
                }
                # Als $sizePropertySuccessfullyUsed $false is, werd 'Size' niet opgevraagd, dus $currentMessageSize blijft $null.

                # Creëer een object met de details van het huidige bericht
                $messageDetail = @{
                    MessageId        = $message.Id
                    Subject          = $message.Subject
                    ReceivedDateTime = $message.ReceivedDateTime
                    Size             = $currentMessageSize # Gebruik de (mogelijk lege) opgehaalde grootte
                    ToRecipients     = $message.ToRecipients | ForEach-Object { $_.EmailAddress.Address } # Sla alleen e-mailadressen op
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
        Write-Host "Indexeren voltooid. $uniqueSenders unieke afzender(s) gevonden."
        
    } catch {
        Write-Error "Fout tijdens het indexeren van de mailbox: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
    # Read-Host "Druk op Enter om terug te keren naar het hoofdmenu" # Verwijderd, Escape in hoofdmenu is de weg terug
}

function Show-SenderOverview {
    param($UserId)

    # Sla huidige consolekleuren op (wordt hersteld door Show-MainMenu bij terugkeer)
    # $originalForegroundColor = $Host.UI.RawUI.ForegroundColor
    # $originalBackgroundColor = $Host.UI.RawUI.BackgroundColor

    # Definieer CGA-kleurenschema (consistent met Show-MainMenu)
    $cgaBgColor = [System.ConsoleColor]::Black
    $cgaFgColor = [System.ConsoleColor]::Green
    $cgaSelectedBgColor = [System.ConsoleColor]::Green
    $cgaSelectedFgColor = [System.ConsoleColor]::Black
    $cgaInstructionFgColor = [System.ConsoleColor]::White
    $cgaWarningFgColor = [System.ConsoleColor]::Red

    if ($null -eq $Script:SenderCache -or $Script:SenderCache.Count -eq 0) {
        $Host.UI.RawUI.ForegroundColor = $cgaWarningFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host
        Write-Host "De mailbox is nog niet geïndexeerd of de index is leeg." -ForegroundColor $cgaWarningFgColor
        Write-Host "Kies optie '1. Indexeer mailbox' in het hoofdmenu om de index op te bouwen." -ForegroundColor $cgaWarningFgColor
        $Host.UI.RawUI.ForegroundColor = $cgaInstructionFgColor
        Write-Host "Druk op Escape of Q om terug te keren." # Aangepast van Read-Host
        while($true){ $key = $Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown); if($key.VirtualKeyCode -eq 27 -or $key.Character.ToString().ToUpper() -eq 'Q'){ break } }
        # Kleuren worden hersteld door Show-MainMenu
        return
    }

    # Converteer de hashtable (nu met domeinen als keys) naar een array voor sortering en weergave
    $domainList = @()
    foreach ($domainKey in $Script:SenderCache.Keys) {
        $domainList += [PSCustomObject]@{
            Domain = $domainKey # De key is het domein
            Name   = $Script:SenderCache[$domainKey].Name # Dit is ook het domein
            Count  = $Script:SenderCache[$domainKey].Count
        }
    }

    $sortedDomains = $domainList | Sort-Object -Property @{Expression="Count"; Descending=$true}, Domain

    if ($sortedDomains.Count -eq 0) {
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host
        Write-Host "Geen domeinen gevonden in de cache."
        $Host.UI.RawUI.ForegroundColor = $cgaInstructionFgColor
        Write-Host "Druk op Escape of Q om terug te keren." # Aangepast van Read-Host
        while($true){ $key = $Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown); if($key.VirtualKeyCode -eq 27 -or $key.Character.ToString().ToUpper() -eq 'Q'){ break } }
        return
    }
    
    $selectedItemIndex = 0
    $menuLoopActive = $true

    while ($menuLoopActive) {
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host

        $title = "Overzicht van afzenderdomeinen voor $UserId"
        $separator = "--------------------------------------------------------------------" # Aangepaste breedte
        $instructionText = "Gebruik ↑/↓, Enter om te kiezen, Esc of Q om terug te keren." # T vervangen door Q
        
        # Bereken dynamische breedte en padding
        $tempMenuContentForWidth = @($title) + @($separator)
        foreach ($domainEntry in $sortedDomains) {
            # Formatteer hoe het item eruit zal zien voor breedteberekening
            $tempMenuContentForWidth += ("{0,-7} {1,-50}" -f $domainEntry.Count, $domainEntry.Domain)
        }
        $tempMenuContentForWidth += $instructionText
        
        $menuWidth = 0
        foreach ($line in $tempMenuContentForWidth) {
            if ($line.Length -gt $menuWidth) { $menuWidth = $line.Length }
        }
        # Voeg een beetje extra toe aan de menuWidth voor de # kolom en selectie
        $menuWidth += 7 # Ruimte voor "#. " en wat speling
        $frameWidth = $menuWidth + 4 
        $consoleWidth = $Host.UI.RawUI.WindowSize.Width
        $leftPaddingSpaces = [Math]::Max(0, ($consoleWidth - $frameWidth) / 2)
        $leftPadding = " " * $leftPaddingSpaces
        $innerFramePadding = "  "

        # Verticale padding
        1..3 | ForEach-Object { Write-Host "" }

        # Teken titel en header
        Write-Host ($leftPadding + $innerFramePadding + $title.PadRight($menuWidth) + $innerFramePadding)
        Write-Host ($leftPadding + $innerFramePadding + $separator.PadRight($menuWidth) + $innerFramePadding)
        $header = "{0,-5} {1,-7} {2,-50}" -f "#", "Aantal", "Domein" # Kolom voor selectienummer
        Write-Host ($leftPadding + $innerFramePadding + $header.PadRight($menuWidth) + $innerFramePadding)
        Write-Host ($leftPadding + $innerFramePadding + $separator.PadRight($menuWidth) + $innerFramePadding)

        # Teken lijst van domeinen
        for ($i = 0; $i -lt $sortedDomains.Count; $i++) {
            $domainEntry = $sortedDomains[$i]
            $itemNumber = $i + 1
            $itemText = "{0,-5} {1,-7} {2,-50}" -f "$itemNumber.", $domainEntry.Count, $domainEntry.Domain
            $lineContent = $innerFramePadding + $itemText.PadRight($menuWidth) + $innerFramePadding
            
            if ($i -eq $selectedItemIndex) {
                Write-Host ($leftPadding + $lineContent) -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor
            } else {
                Write-Host ($leftPadding + $lineContent) -ForegroundColor $cgaFgColor -BackgroundColor $cgaBgColor
            }
        }
        
        Write-Host ($leftPadding + $innerFramePadding + $separator.PadRight($menuWidth) + $innerFramePadding)
        Write-Host ($leftPadding + $innerFramePadding + $instructionText.PadRight($menuWidth) + $innerFramePadding) -ForegroundColor $cgaInstructionFgColor

        # Wacht op toetsaanslag
        $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)
        $choiceToProcess = $null

        switch ($keyInfo.VirtualKeyCode) {
            38 { # UpArrow
                $selectedItemIndex--
                if ($selectedItemIndex -lt 0) { $selectedItemIndex = $sortedDomains.Count - 1 }
            }
            40 { # DownArrow
                $selectedItemIndex++
                if ($selectedItemIndex -ge $sortedDomains.Count) { $selectedItemIndex = 0 }
            }
            13 { # Enter
                $choiceToProcess = $selectedItemIndex # Gebruik index direct
            }
            27 { # Escape
                $menuLoopActive = $false # Stop de lus, keer terug naar hoofdmenu
            }
            default {
                $charPressed = $keyInfo.Character.ToString().ToUpper()
                if ($charPressed -eq 'Q') { # T vervangen door Q
                    $menuLoopActive = $false # Stop de lus, keer terug
                } elseif ($charPressed -match "^\d+$") { # Als een nummer is ingevoerd
                    $numChoice = [int]$charPressed
                    if ($numChoice -ge 1 -and $numChoice -le $sortedDomains.Count) {
                        $selectedItemIndex = $numChoice - 1
                        $choiceToProcess = $selectedItemIndex
                    }
                }
            }
        }

        if (-not $menuLoopActive) { # Als Esc of T is gedrukt
            # Kleuren worden hersteld door Show-MainMenu
            return
        }

        if ($choiceToProcess -ne $null) {
            $selectedDomainInfo = $sortedDomains[$choiceToProcess] # $choiceToProcess is de index
            
            # Herstel standaard consolekleuren voordat een subfunctie wordt aangeroepen
            # Dit wordt afgehandeld door Show-MainMenu bij terugkeer, en subfuncties moeten hun eigen kleuren instellen indien nodig.
            # Voor nu, roepen we de volgende functie aan. Deze moet ook CGA-bewust zijn of kleuren herstellen.
            
            # BELANGRIJK: Show-EmailsFromSelectedSender moet worden aangepast om een domein te accepteren
            # en de $SenderInfo parameter moet mogelijk hernoemd worden naar $DomainInfo of iets dergelijks.
            # Voor nu, passen we de aanroep aan met de verwachting dat $SenderInfo een object is met een .Domain eigenschap.
            # De naam $selectedSenderInfo wordt $selectedDomainInfo.
            # De parameter voor Show-EmailsFromSelectedSender heet $SenderInfo, we sturen een object met .Email, .Name, .Count
            # We moeten een object maken dat overeenkomt, of Show-EmailsFromSelectedSender aanpassen.
            # Laten we $selectedDomainInfo direct doorgeven en Show-EmailsFromSelectedSender aanpassen.
            Show-EmailsFromSelectedSender -UserId $UserId -SenderInfo $selectedDomainInfo # $selectedDomainInfo heeft .Domain, .Name, .Count
            
            # Na terugkeer van Show-EmailsFromSelectedSender, wordt de lijst opnieuw getoond.
            # De cache kan gewijzigd zijn, dus de $sortedDomains moeten mogelijk opnieuw worden opgebouwd.
            # De huidige lus zal dit automatisch doen bij de volgende iteratie als $menuLoopActive nog true is.
            # Echter, als Show-EmailsFromSelectedSender terugkeert, willen we meestal terug naar het hoofdmenu
            # om de gebruiker niet vast te houden in een diepere menustructuur zonder duidelijke "terug" optie.
            # De huidige Show-EmailsFromSelectedSender retourneert naar Show-SenderOverview, die dan opnieuw de lijst toont.
            # Dit is acceptabel. We moeten wel de $sortedDomains opnieuw laden als de cache is veranderd.
            # Dit gebeurt aan het begin van de Show-SenderOverview functie.
            # Dus, na een actie in Show-EmailsFromSelectedSender, zal deze functie opnieuw de data laden.
            # We moeten de lus niet stoppen, tenzij de gebruiker expliciet terug wil (Esc/T).
            # De $sortedDomains worden aan het begin van de functie geladen, niet in de lus. Dit moet veranderen.
            # Herlaad de data hier als we in de lus blijven na een actie.
            # Voor nu, na een actie in Show-EmailsFromSelectedSender, keert die functie terug,
            # en deze lus gaat verder, en tekent opnieuw. De data moet dan wel actueel zijn.
            # De eenvoudigste manier is om Show-SenderOverview opnieuw aan te roepen of de data hier te herladen.
            # Laten we de data herladen:
            $domainList = @()
            foreach ($domainKey in $Script:SenderCache.Keys) {
                $domainList += [PSCustomObject]@{ Domain = $domainKey; Name = $Script:SenderCache[$domainKey].Name; Count  = $Script:SenderCache[$domainKey].Count }
            }
            $sortedDomains = $domainList | Sort-Object -Property @{Expression="Count"; Descending=$true}, Domain
            if ($sortedDomains.Count -eq 0 -or $selectedItemIndex -ge $sortedDomains.Count) { # Als alle items van een domein zijn verwijderd
                $selectedItemIndex = 0 # Reset selectie
                if ($sortedDomains.Count -eq 0) { # Als er helemaal geen domeinen meer zijn
                    $menuLoopActive = $false # Verlaat de functie
                    # De melding "Geen domeinen gevonden" wordt dan getoond bij de volgende aanroep van Show-SenderOverview
                }
            }
            # Ga door met de volgende iteratie van de lus om het menu opnieuw te tekenen
        }
    }
    # Kleuren worden hersteld door Show-MainMenu
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
}

# Nieuwe functie om e-mails van een geselecteerde afzender te tonen en acties te starten
function Show-EmailsFromSelectedSender {
    param (
        [string]$UserId,
        [PSCustomObject]$SenderInfo # Ontvangt nu een object met .Domain, .Name (is domein), .Count
    )

    $domainName = $SenderInfo.Domain # Gebruik .Domain ipv .Email
    $normalizedDomainKey = $domainName.ToLowerInvariant()

    # Blijf in een lus zolang er berichten zijn voor dit domein en de gebruiker niet terug wil
    while ($Script:SenderCache.ContainsKey($normalizedDomainKey) -and $Script:SenderCache[$normalizedDomainKey].Messages.Count -gt 0) {
        # CGA Kleuren moeten hier ook worden ingesteld als deze functie direct wordt aangeroepen
        # en niet via een menu dat al kleuren beheert.
        # CGA Kleuren
        $cgaBgColor = [System.ConsoleColor]::Black
        $cgaFgColor = [System.ConsoleColor]::Green
        $cgaSelectedBgColor = [System.ConsoleColor]::Green
        $cgaSelectedFgColor = [System.ConsoleColor]::Black
        $cgaInstructionFgColor = [System.ConsoleColor]::White

        $selectedEmailIndex = 0
        $selectedActionIndex = 0 # Voor het actiemenu onderaan
        $spaceSelectedMessageIds = [System.Collections.Generic.HashSet[string]]::new() # Voor spatiebalkselectie

        # Acties die onderaan de e-maillijst getoond worden
        $bottomMenuItems = @(
            "Acties op selectie (Enter/Spatie)",  # Index 0 - Aangepaste tekst
            "Beheer ALLE e-mails van dit domein", # Index 1
            "Terug naar domeinoverzicht (Esc/Q)"  # Index 2
        )
        $currentFocusIsEmailList = $true # True als focus op e-maillijst, False als op actiemenu

        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host 
        
        $cachedDomainEntry = $Script:SenderCache[$normalizedDomainKey]
        # Sorteer berichten hier eenmalig, tenzij de lijst verandert
        $messagesFromDomain = $cachedDomainEntry.Messages | Sort-Object ReceivedDateTime -Descending
        
        if ($messagesFromDomain.Count -eq 0) {
            Write-Host "Geen e-mails (meer) in de cache voor domein '$domainName'." -ForegroundColor $cgaInstructionFgColor
            Write-Host "Druk op Escape of Q om terug te keren." -ForegroundColor $cgaInstructionFgColor
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            while($true){ $key = $Host.UI.RawUI.ReadKey($readKeyOptions); if($key.VirtualKeyCode -eq 27 -or $key.Character.ToString().ToUpper() -eq 'Q'){ break } }
            return
        }

        # Hoofd lus voor dit menu
        $emailMenuLoopActive = $true
        while ($emailMenuLoopActive) {
            $Host.UI.RawUI.ForegroundColor = $cgaFgColor
            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
            Clear-Host

            Write-Host "E-mails van domein: $($cachedDomainEntry.Name)"
            Write-Host "Aantal in cache: $($cachedDomainEntry.Count)"
            Write-Host "Gebruik ↑/↓ om te navigeren, Enter om te selecteren/lezen, Tab om focus te wisselen, Esc/Q om terug te keren." -ForegroundColor $cgaInstructionFgColor
            Write-Host "-------------------------------------------------------------------------------------------------------------------"
            Write-Host ("{0,-5} {1,-60} {2,-20} {3,-15}" -f "#", "Onderwerp", "Ontvangen Op", "Grootte (Bytes)")
            Write-Host "-------------------------------------------------------------------------------------------------------------------"

            # Toon e-maillijst
            for ($i = 0; $i -lt $messagesFromDomain.Count; $i++) {
                $message = $messagesFromDomain[$i]
                $itemNumber = $i + 1
                $subjectDisplay = if ($message.Subject) { ($message.Subject | Select-Object -First 1) } else { "(Geen onderwerp)" }
                if ($subjectDisplay.Length -gt 57) { $subjectDisplay = $subjectDisplay.Substring(0, 57) + "..." }
                $receivedDisplay = if ($message.ReceivedDateTime) { Get-Date $message.ReceivedDateTime -Format "yyyy-MM-dd HH:mm" } else { "N/B" }
                $sizeDisplay = if ($message.Size -ne $null) { $message.Size } else { "N/B" }
                
                $selectionPrefix = "   " # Drie spaties voor niet-geselecteerd
                if ($spaceSelectedMessageIds.Contains($message.MessageId)) {
                    $selectionPrefix = "[*]" # Visuele indicator voor spatie-selectie
                }
                $lineText = "{0} {1,-5} {2,-57} {3,-20} {4,-15}" -f $selectionPrefix, "$itemNumber.", $subjectDisplay, $receivedDisplay, $sizeDisplay
                
                if ($currentFocusIsEmailList -and $i -eq $selectedEmailIndex) {
                    Write-Host $lineText -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor
                } else {
                    Write-Host $lineText
                }
            }
            Write-Host "-------------------------------------------------------------------------------------------------------------------"
            Write-Host ("Geselecteerd met spatie: {0}" -f $spaceSelectedMessageIds.Count) -ForegroundColor $cgaInstructionFgColor
            # Toon actiemenu onderaan
            for ($i = 0; $i -lt $bottomMenuItems.Count; $i++) {
                $actionText = $bottomMenuItems[$i]
                if (-not $currentFocusIsEmailList -and $i -eq $selectedActionIndex) {
                    Write-Host "> $($actionText)" -ForegroundColor $cgaSelectedFgColor -BackgroundColor $cgaSelectedBgColor
                } else {
                    Write-Host "  $($actionText)"
                }
            }
            Write-Host "-------------------------------------------------------------------------------------------------------------------"

            # Wacht op toetsaanslag
            $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
            $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)

            if ($currentFocusIsEmailList) {
                switch ($keyInfo.VirtualKeyCode) {
                    38 { # UpArrow
                        $selectedEmailIndex--
                        if ($selectedEmailIndex -lt 0) { $selectedEmailIndex = $messagesFromDomain.Count - 1 }
                    }
                    40 { # DownArrow
                        $selectedEmailIndex++
                        if ($selectedEmailIndex -ge $messagesFromDomain.Count) { $selectedEmailIndex = 0 }
                    }
                    13 { # Enter - Bekijk geselecteerde e-mail
                        if ($messagesFromDomain.Count -gt 0 -and $selectedEmailIndex -lt $messagesFromDomain.Count) {
                            $Host.UI.RawUI.ForegroundColor = $cgaFgColor # Herstel kleuren voor subfunctie
                            $Host.UI.RawUI.BackgroundColor = $cgaBgColor
                            Show-EmailBody -UserId $UserId -MessageObject $messagesFromDomain[$selectedEmailIndex]
                            # Na terugkeer, herlaad de lijst niet, ga gewoon door met de lus om opnieuw te tekenen
                        }
                    }
                    32 { # Spacebar - Toggle selectie van huidige item
                        if ($messagesFromDomain.Count -gt 0 -and $selectedEmailIndex -lt $messagesFromDomain.Count) {
                            $currentMessageId = $messagesFromDomain[$selectedEmailIndex].MessageId
                            if ($spaceSelectedMessageIds.Contains($currentMessageId)) {
                                $spaceSelectedMessageIds.Remove($currentMessageId) | Out-Null
                            } else {
                                $spaceSelectedMessageIds.Add($currentMessageId) | Out-Null
                            }
                        }
                    }
                    9 { # Tab - Wissel focus naar actiemenu
                        $currentFocusIsEmailList = $false
                        $selectedActionIndex = 0 # Reset selectie in actiemenu
                    }
                    27 { $emailMenuLoopActive = $false } # Escape
                }
                if ($keyInfo.Character.ToString().ToUpper() -eq 'Q') { $emailMenuLoopActive = $false }

            } else { # Focus is op actiemenu
                switch ($keyInfo.VirtualKeyCode) {
                    38 { # UpArrow
                        $selectedActionIndex--
                        if ($selectedActionIndex -lt 0) { $selectedActionIndex = $bottomMenuItems.Count - 1 }
                    }
                    40 { # DownArrow
                        $selectedActionIndex++
                        if ($selectedActionIndex -ge $bottomMenuItems.Count) { $selectedActionIndex = 0 }
                    }
                    13 { # Enter - Voer geselecteerde actie uit
                        $chosenAction = $bottomMenuItems[$selectedActionIndex]
                        if ($chosenAction -like "Acties op selectie*") {
                            $messagesForAction = New-Object System.Collections.Generic.List[PSObject]
                            if ($spaceSelectedMessageIds.Count -gt 0) {
                                # Gebruik met spatie geselecteerde items
                                $messagesFromDomain | Where-Object { $spaceSelectedMessageIds.Contains($_.MessageId) } | ForEach-Object { $messagesForAction.Add($_) }
                            } elseif ($messagesFromDomain.Count -gt 0 -and $selectedEmailIndex -lt $messagesFromDomain.Count) {
                                # Gebruik huidig gehighlight item als er geen spatie-selecties zijn
                                $messagesForAction.Add($messagesFromDomain[$selectedEmailIndex])
                            }

                            if ($messagesForAction.Count -gt 0) {
                                $Host.UI.RawUI.ForegroundColor = $cgaFgColor; $Host.UI.RawUI.BackgroundColor = $cgaBgColor
                                # Roep een nieuwe functie aan die een lijst van berichten kan verwerken
                                Perform-ActionOnMultipleEmails -UserId $UserId -MessagesToProcess $messagesForAction -DomainToUpdateCache $domainName
                                $spaceSelectedMessageIds.Clear() # Wis spatie-selectie na actie
                                # Herlaad berichtenlijst na actie
                                if ($Script:SenderCache.ContainsKey($normalizedDomainKey)) {
                                     $messagesFromDomain = $Script:SenderCache[$normalizedDomainKey].Messages | Sort-Object ReceivedDateTime -Descending
                                     if ($selectedEmailIndex -ge $messagesFromDomain.Count) {$selectedEmailIndex = [Math]::Max(0, $messagesFromDomain.Count - 1)}
                                     if ($messagesFromDomain.Count -eq 0) { $emailMenuLoopActive = $false }
                                } else { 
                                    $emailMenuLoopActive = $false 
                                }
                            } else {
                                Write-Warning "Geen e-mails geselecteerd voor actie."
                                # Wacht even zodat de gebruiker de waarschuwing kan lezen
                                Start-Sleep -Seconds 2 
                            }
                        } elseif ($chosenAction -like "Beheer ALLE e-mails*") {
                            $Host.UI.RawUI.ForegroundColor = $cgaFgColor; $Host.UI.RawUI.BackgroundColor = $cgaBgColor
                            $allMessagesWereModified = Perform-ActionOnAllSenderEmails -UserId $UserId -SenderDomain $domainName -AllMessages $messagesFromDomain
                            if ($allMessagesWereModified -or -not $Script:SenderCache.ContainsKey($normalizedDomainKey)) {
                                $emailMenuLoopActive = $false # Verlaat dit menu als domein weg is
                            } else {
                                # Herlaad berichtenlijst
                                $messagesFromDomain = $Script:SenderCache[$normalizedDomainKey].Messages | Sort-Object ReceivedDateTime -Descending
                                if ($selectedEmailIndex -ge $messagesFromDomain.Count) {$selectedEmailIndex = [Math]::Max(0, $messagesFromDomain.Count - 1)}
                                if ($messagesFromDomain.Count -eq 0) { $emailMenuLoopActive = $false } # Verlaat als er geen berichten meer zijn
                            }
                        } elseif ($chosenAction -like "Terug*") {
                            $emailMenuLoopActive = $false
                        }
                    }
                    9 { # Tab - Wissel focus naar e-maillijst
                        $currentFocusIsEmailList = $true
                        # $selectedEmailIndex blijft behouden
                    }
                    27 { $emailMenuLoopActive = $false } # Escape
                }
                 if ($keyInfo.Character.ToString().ToUpper() -eq 'Q' -and $bottomMenuItems[$selectedActionIndex] -like "Terug*") { $emailMenuLoopActive = $false }
            }
             # Als er geen berichten meer zijn na een actie, verlaat de lus
            if ($emailMenuLoopActive) { # Alleen controleren als we niet al besloten hebben te stoppen
                if ($Script:SenderCache.ContainsKey($normalizedDomainKey)) {
                    if ($Script:SenderCache[$normalizedDomainKey].Messages.Count -eq 0) { $emailMenuLoopActive = $false }
                } else {
                    $emailMenuLoopActive = $false # Domein is verwijderd
                }
            }
        } # Einde $emailMenuLoopActive while
    } # Einde hoofd while ($Script:SenderCache.ContainsKey...

    # Als de lus eindigt omdat het domein geen berichten meer heeft of niet meer in de cache is:
    # De aanroepende functie (Show-SenderOverview) zal de UI verder afhandelen.
}

# Nieuwe functie voor acties op meerdere (met spatie) geselecteerde e-mails
function Perform-ActionOnMultipleEmails {
    param (
        [string]$UserId,
        [System.Collections.Generic.List[PSObject]]$MessagesToProcess,
        [string]$DomainToUpdateCache
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

    if ($actionToExecute) {
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor; $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host

        if ($actionToExecute -like "1. Verwijder*") {
            if (Get-Confirmation -PromptMessage "Weet u zeker dat u deze $($MessagesToProcess.Count) e-mail(s) permanent wilt verwijderen?") {
                Write-Host "Starten met verwijderen van $($MessagesToProcess.Count) e-mail(s)..."
                $processedCount = 0; $errorCount = 0
                foreach ($message in $MessagesToProcess) {
                    $processedCount++
                    Write-Progress -Activity "Geselecteerde e-mails verwijderen" -Status "Verwijderen: $($message.Subject)" -PercentComplete (($processedCount / $MessagesToProcess.Count) * 100)
                    try {
                        Remove-MgUserMessage -UserId $UserId -MessageId $message.MessageId -ErrorAction Stop
                        Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $message.MessageId
                    } catch {
                        Write-Warning "Fout bij verwijderen e-mail ID $($message.MessageId): $($_.Exception.Message)"
                        $errorCount++
                    }
                }
                Write-Progress -Activity "Geselecteerde e-mails verwijderen" -Completed
                Write-Host "Verwijderen voltooid. $($MessagesToProcess.Count - $errorCount) e-mail(s) verwijderd."
                if ($errorCount -gt 0) { Write-Warning "$errorCount e-mail(s) konden niet worden verwijderd." }
            } else { Write-Host "Verwijderen geannuleerd." }
            Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
            while($Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown).VirtualKeyCode -ne 27) {}

        } elseif ($actionToExecute -like "2. Verplaats*") {
            $destinationFolderId = Get-MailFolderSelection -UserId $UserId
            if ($destinationFolderId) {
                $destinationFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $destinationFolderId -ErrorAction SilentlyContinue
                if (Get-Confirmation -PromptMessage "Weet u zeker dat u deze $($MessagesToProcess.Count) e-mail(s) wilt verplaatsen naar '$($destinationFolder.DisplayName)'?") {
                    Write-Host "Starten met verplaatsen van $($MessagesToProcess.Count) e-mail(s) naar '$($destinationFolder.DisplayName)'..."
                    $processedCount = 0; $errorCount = 0
                    foreach ($message in $MessagesToProcess) {
                        $processedCount++
                        Write-Progress -Activity "Geselecteerde e-mails verplaatsen" -Status "Verplaatsen: $($message.Subject)" -PercentComplete (($processedCount / $MessagesToProcess.Count) * 100)
                        try {
                            Move-MgUserMessage -UserId $UserId -MessageId $message.MessageId -DestinationId $destinationFolderId -ErrorAction Stop
                            Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $message.MessageId
                        } catch {
                            Write-Warning "Fout bij verplaatsen e-mail ID $($message.MessageId): $($_.Exception.Message)"
                            $errorCount++
                        }
                    }
                    Write-Progress -Activity "Geselecteerde e-mails verplaatsen" -Completed
                    Write-Host "Verplaatsen voltooid. $($MessagesToProcess.Count - $errorCount) e-mail(s) verplaatst."
                    if ($errorCount -gt 0) { Write-Warning "$errorCount e-mail(s) konden niet worden verplaatst." }
                } else { Write-Host "Verplaatsen geannuleerd." }
            } else { Write-Host "Verplaatsen geannuleerd (geen doelmap geselecteerd)." }
            Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor
            while($Host.UI.RawUI.ReadKey([System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown).VirtualKeyCode -ne 27) {}
        } elseif ($actionToExecute -like "3. Terug*") {
            # Do nothing, function will return
        }
    }
}

# Helper functie om de body van een e-mail te tonen
function Show-EmailBody {
    param (
        [string]$UserId,
        [PSCustomObject]$MessageObject # Het volledige $messageDetail object uit de cache of direct van Graph
    )

    # CGA Kleuren (ervan uitgaande dat de aanroeper de kleuren beheert voor/na deze functie)
    $cgaBgColor = [System.ConsoleColor]::Black
    $cgaFgColor = [System.ConsoleColor]::Green
    $cgaInstructionFgColor = [System.ConsoleColor]::White
    
    $Host.UI.RawUI.ForegroundColor = $cgaFgColor
    $Host.UI.RawUI.BackgroundColor = $cgaBgColor
    Clear-Host

    Write-Host "Volledige body van e-mail:"
    Write-Host "Onderwerp    : $($MessageObject.Subject)"
    Write-Host "Ontvangen op : $(Get-Date $MessageObject.ReceivedDateTime -Format "yyyy-MM-dd HH:mm:ss")"
    Write-Host "ID           : $($MessageObject.MessageId)"
    Write-Host "----------------------------------------------------"

    # Haal de volledige body op als die nog niet in $MessageObject zit (bijv. als het uit de cache komt zonder body)
    $bodyContent = $MessageObject.Body # Als 'Body' al een property is met .Content
    if ($MessageObject.PSObject.Properties["Body"] -and $MessageObject.Body.PSObject.Properties["Content"]) {
        $bodyContent = $MessageObject.Body.Content
    } elseif ($MessageObject.PSObject.Properties["BodyPreview"]) {
        # Fallback naar BodyPreview als volledige body niet direct beschikbaar is
        # Idealiter zou je hier de volledige body ophalen via Get-MgUserMessage als die ontbreekt
        Write-Warning "Volledige body niet direct beschikbaar in het object, toon BodyPreview."
        $bodyContent = $MessageObject.BodyPreview
    } else {
         # Probeer de volledige body op te halen als deze ontbreekt
        try {
            Write-Host "Ophalen van volledige body van server..."
            $fullMessage = Get-MgUserMessage -UserId $UserId -MessageId $MessageObject.MessageId -Property "body" -ErrorAction Stop
            if ($fullMessage -and $fullMessage.Body) {
                $bodyContent = $fullMessage.Body.Content
                if ($fullMessage.Body.ContentType -eq "html") {
                    Write-Warning "De body is in HTML-formaat. HTML-tags worden als platte tekst weergegeven."
                }
            } else {
                $bodyContent = "(Kon volledige body niet ophalen)"
            }
        } catch {
            Write-Error "Fout bij ophalen volledige body: $($_.Exception.Message)"
            $bodyContent = "(Fout bij ophalen body)"
        }
    }
    
    Write-Host $bodyContent
    Write-Host "----------------------------------------------------"
    Write-Host "Druk op Escape om terug te keren." -ForegroundColor $cgaInstructionFgColor

    while ($true) {
        $readKeyOptions = [System.Management.Automation.Host.ReadKeyOptions]::NoEcho -bor [System.Management.Automation.Host.ReadKeyOptions]::IncludeKeyDown
        $keyInfo = $Host.UI.RawUI.ReadKey($readKeyOptions)
        if ($keyInfo.VirtualKeyCode -eq 27) { # Escape
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
    Write-Host "Geselecteerde e-mail:"
    Write-Host "Onderwerp : $($MessageObject.Subject)"
    Write-Host "Ontvangen: $($MessageObject.ReceivedDateTime)"
    Write-Host "ID        : $($MessageObject.MessageId)"
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
    $emailIdDisplay = $MessageObject.MessageId

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
                        Write-Host "Verwijderen van e-mail ID $($MessageObject.MessageId)..."
                    Remove-MgUserMessage -UserId $UserId -MessageId $MessageObject.MessageId -ErrorAction Stop
                    Write-Host "E-mail succesvol verwijderd van server."
                    # Update cache
                    Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $MessageObject.MessageId
                } catch {
                        Write-Error "Fout bij het verwijderen van e-mail ID $($MessageObject.MessageId): $($_.Exception.Message)"
                    }
                } else { Write-Host "Verwijderen geannuleerd." }
                # Read-Host "Druk op Enter om door te gaan." # Verwijderd
            } elseif ($actionToExecute -like "2. Verplaats*") {
                $destinationFolderId = Get-MailFolderSelection -UserId $UserId # Deze moet ook interactief worden
                if ($destinationFolderId) {
                    $destinationFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $destinationFolderId -ErrorAction SilentlyContinue
                    if (Get-Confirmation -PromptMessage "Weet u zeker dat u deze e-mail wilt verplaatsen naar '$($destinationFolder.DisplayName)'?") {
                        try {
                            Write-Host "Verplaatsen van e-mail ID $($MessageObject.MessageId) naar '$($destinationFolder.DisplayName)'..."
                            Move-MgUserMessage -UserId $UserId -MessageId $MessageObject.MessageId -DestinationId $destinationFolderId -ErrorAction Stop
                            Write-Host "E-mail succesvol verplaatst."
                        # Update cache
                            Update-SenderCache -DomainToUpdate $DomainToUpdateCache -MessageIdToRemove $MessageObject.MessageId
                        } catch {
                            Write-Error "Fout bij het verplaatsen van e-mail ID $($MessageObject.MessageId): $($_.Exception.Message)"
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
    [OutputType([bool])] # OutputType was hier al, maar stond verkeerd in eerdere diff. Nu correct geplaatst.
    param (
        [string]$UserId,
        [string]$SenderDomain, # Aangepast van SenderEmail naar SenderDomain
        [System.Collections.Generic.List[PSObject]]$AllMessages # Dit zijn alle berichten van het domein
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
        
        # Wis de vorige menu-opties (simpele aanpak: overschrijf met lege regels)
        # Dit is lastig zonder precieze cursor controle. Voor nu, Clear-Host aan begin van lus.
        # Clear-Host was al aan het begin van de functie.
        # We moeten de opties overschrijven.
        # Voor nu, de Clear-Host aan het begin van de functie zal het scherm leegmaken,
        # en de lus zal het menu opnieuw tekenen.

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
    Clear-Host
    Write-Host "Beheer e-mails van specifieke afzender voor $UserId"
    Write-Host "----------------------------------------------------"

    if ($null -eq $Script:SenderCache -or $Script:SenderCache.Count -eq 0) {
        Write-Warning "De mailbox is nog niet geïndexeerd of de index is leeg."
        Write-Warning "Kies optie '1. Indexeer mailbox' in het hoofdmenu om de index op te bouwen."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    $senderEmail = Read-Host "Voer het e-mailadres in van de afzender"
    if (-not $Script:SenderCache.ContainsKey($senderEmail.ToLowerInvariant())) {
        Write-Warning "Afzender '$senderEmail' niet gevonden in de cache. Controleer het e-mailadres of indexeer de mailbox opnieuw."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    $senderInfo = $Script:SenderCache[$senderEmail.ToLowerInvariant()]
    Write-Host "Afzender geselecteerd: $($senderInfo.Name) <$senderEmail>"
    Write-Host "Aantal e-mails in cache: $($senderInfo.Count)"
    Write-Host ""
    Write-Host "Kies een actie:"
    Write-Host "1. Verwijder alle e-mails van deze afzender"
    Write-Host "2. Verplaats alle e-mails van deze afzender naar een andere map"
    Write-Host "3. Terug naar hoofdmenu"
    
    $actionChoice = Read-Host "Kies een optie (1-3)"

    switch ($actionChoice) {
        "1" { Delete-MailsFromSender -UserId $UserId -SenderEmail $senderEmail }
        "2" { Move-MailsFromSender -UserId $UserId -SenderEmail $senderEmail }
        "3" { return } # Terug naar hoofdmenu
        default {
            Write-Warning "Ongeldige keuze."
        }
    }
    Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
}

function Delete-MailsFromSender {
    param(
        [string]$UserId,
        [string]$SenderEmail
    )
    Clear-Host
    Write-Host "Verwijderen van e-mails van: $SenderEmail"
    Write-Host "-------------------------------------------"

    $confirmation = Read-Host "WAARSCHUWING: Weet u zeker dat u ALLE e-mails van '$SenderEmail' permanent wilt verwijderen? (ja/nee)"
    if ($confirmation -ne 'ja') {
        Write-Host "Verwijderen geannuleerd."
        return
    }

    try {
        Write-Host "Zoeken naar e-mails van '$SenderEmail'..."
        # Filter op het e-mailadres van de afzender in het 'From' veld.
        # De -All parameter zorgt ervoor dat alle overeenkomende berichten worden opgehaald.
        $messagesToDelete = Get-MgUserMessage -UserId $UserId -Filter "from/emailAddress/address eq '$SenderEmail'" -All -ErrorAction Stop
        
        if ($null -eq $messagesToDelete -or $messagesToDelete.Count -eq 0) {
            Write-Host "Geen e-mails gevonden van '$SenderEmail'."
            return
        }

        $count = $messagesToDelete.Count
        Write-Host "$count e-mail(s) gevonden van '$SenderEmail'. Starten met verwijderen..."
        
        $deletedCount = 0
        $errorCount = 0

        foreach ($message in $messagesToDelete) {
            try {
                Write-Progress -Activity "E-mails verwijderen" -Status "Verwijderen: $($message.Subject)" -PercentComplete (($deletedCount / $count) * 100)
                Remove-MgUserMessage -UserId $UserId -MessageId $message.Id -ErrorAction Stop
                $deletedCount++
            } catch {
                Write-Warning "Fout bij het verwijderen van e-mail met ID $($message.Id) (Onderwerp: $($message.Subject)): $($_.Exception.Message)"
                $errorCount++
            }
        }
        Write-Progress -Activity "E-mails verwijderen" -Completed

        Write-Host "Verwijderen voltooid."
        Write-Host "$deletedCount e-mail(s) succesvol verwijderd."
        if ($errorCount -gt 0) {
            Write-Warning "$errorCount e-mail(s) konden niet worden verwijderd."
        }

        Write-Warning "De lokale cache is mogelijk niet meer accuraat. Het wordt aanbevolen de mailbox opnieuw te indexeren."

    } catch {
        Write-Error "Fout tijdens het zoeken of verwijderen van e-mails: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
}

function Move-MailsFromSender {
    param(
        [string]$UserId,
        [string]$SenderEmail
    )
    Clear-Host
    Write-Host "Verplaatsen van e-mails van: $SenderEmail"
    Write-Host "--------------------------------------------"

    # Stap 1: Selecteer doelmap
    $destinationFolderId = Get-MailFolderSelection -UserId $UserId
    if (-not $destinationFolderId) {
        Write-Host "Verplaatsen geannuleerd (geen doelmap geselecteerd)."
        return
    }

    $destinationFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $destinationFolderId -ErrorAction SilentlyContinue
    Write-Host "Geselecteerde doelmap: $($destinationFolder.DisplayName)"
    Write-Host ""

    $confirmation = Read-Host "WAARSCHUWING: Weet u zeker dat u ALLE e-mails van '$SenderEmail' wilt verplaatsen naar '$($destinationFolder.DisplayName)'? (ja/nee)"
    if ($confirmation -ne 'ja') {
        Write-Host "Verplaatsen geannuleerd."
        return
    }

    try {
        Write-Host "Zoeken naar e-mails van '$SenderEmail'..."
        $messagesToMove = Get-MgUserMessage -UserId $UserId -Filter "from/emailAddress/address eq '$SenderEmail'" -All -ErrorAction Stop
        
        if ($null -eq $messagesToMove -or $messagesToMove.Count -eq 0) {
            Write-Host "Geen e-mails gevonden van '$SenderEmail'."
            return
        }

        $count = $messagesToMove.Count
        Write-Host "$count e-mail(s) gevonden van '$SenderEmail'. Starten met verplaatsen naar '$($destinationFolder.DisplayName)'..."
        
        $movedCount = 0
        $errorCount = 0

        foreach ($message in $messagesToMove) {
            try {
                Write-Progress -Activity "E-mails verplaatsen" -Status "Verplaatsen: $($message.Subject) naar $($destinationFolder.DisplayName)" -PercentComplete (($movedCount / $count) * 100)
                Move-MgUserMessage -UserId $UserId -MessageId $message.Id -DestinationId $destinationFolderId -ErrorAction Stop
                $movedCount++
            } catch {
                Write-Warning "Fout bij het verplaatsen van e-mail met ID $($message.Id) (Onderwerp: $($message.Subject)): $($_.Exception.Message)"
                $errorCount++
            }
        }
        Write-Progress -Activity "E-mails verplaatsen" -Completed

        Write-Host "Verplaatsen voltooid."
        Write-Host "$movedCount e-mail(s) succesvol verplaatst naar '$($destinationFolder.DisplayName)'."
        if ($errorCount -gt 0) {
            Write-Warning "$errorCount e-mail(s) konden niet worden verplaatst."
        }

        Write-Warning "De lokale cache is mogelijk niet meer accuraat. Het wordt aanbevolen de mailbox opnieuw te indexeren."

    } catch {
        Write-Error "Fout tijdens het zoeken of verplaatsen van e-mails: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
}

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
    
    $searchTerm = Read-Host "Voer zoekterm in (zoekt in onderwerp, body, afzender)"
    if ([string]::IsNullOrWhiteSpace($searchTerm)) {
        Write-Warning "Geen zoekterm ingevoerd. Zoekactie geannuleerd."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    try {
        Write-Host "Zoeken naar e-mails met term: '$searchTerm'..."
        # De -Search parameter gebruikt de Microsoft Search KQL syntax.
        # Standaard zoekt het in meerdere velden zoals onderwerp, body, afzender.
        # We selecteren specifieke properties voor een snellere en relevantere output.
        
        $getMgUserMessageParams = @{
            UserId = $UserId
            Search = $searchTerm
            Top = 100 # Beperk het aantal resultaten
            Property = "subject,from,receivedDateTime,hasAttachments"
            ErrorAction = "Stop"
        }

        if ($IsTestRun.IsPresent) {
            $getMgUserMessageParams.OrderBy = "receivedDateTime desc"
            Write-Host "(Testmodus actief: resultaten gesorteerd op nieuwste eerst)"
        }

        $foundMessages = Get-MgUserMessage @getMgUserMessageParams
        
        if ($null -eq $foundMessages -or $foundMessages.Count -eq 0) {
            Write-Host "Geen e-mails gevonden die overeenkomen met de zoekterm '$searchTerm'."
        } else {
            Write-Host "$($foundMessages.Count) e-mail(s) gevonden. Selecteer een e-mail voor acties:"
            Write-Host "----------------------------------------------------------------------------------------------------"
            
            $selectableMessages = @{}
            $i = 1
            foreach ($message in $foundMessages) {
                $fromAddress = if ($message.From -and $message.From.EmailAddress) { $message.From.EmailAddress.Address } else { "N/B" }
                $subject = if ($message.Subject) { $message.Subject } else { "(Geen onderwerp)" }
                $received = if ($message.ReceivedDateTime) { Get-Date $message.ReceivedDateTime -Format "yyyy-MM-dd HH:mm" } else { "N/B" }
                
                Write-Host ("{0}. Onderwerp    : {1}" -f $i, $subject)
                Write-Host ("   Van          : {0}" -f $fromAddress)
                Write-Host ("   Ontvangen op : {0}" -f $received)
                Write-Host ("   ID           : {0}" -f $message.Id)
                Write-Host "----------------------------------------------------------------------------------------------------"
                $selectableMessages[$i] = $message.Id
                $i++
            }

            Write-Host "T. Terug naar hoofdmenu"
            $choice = Read-Host "Kies een e-mailnummer (1-$($i-1)) of T om terug te keren"

            if ($choice -eq 'T' -or $choice -eq 't') {
                # Doe niets, keer terug naar hoofdmenu via de Read-Host aan het einde van de functie
            } elseif ($selectableMessages.ContainsKey($choice)) {
                $selectedMessageId = $selectableMessages[$choice]
                Show-EmailActionsMenu -UserId $UserId -MessageId $selectedMessageId
                # Na Show-EmailActionsMenu, keer terug naar hoofdmenu, dus geen extra Read-Host hier nodig.
                return # Keer direct terug om de Read-Host aan het einde van Search-Mail te vermijden
            } else {
                Write-Warning "Ongeldige keuze."
            }
        }
    } catch {
        Write-Error "Fout tijdens het zoeken naar e-mails: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
    
    Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
}

function Show-EmailActionsMenu {
    param(
        [string]$UserId,
        [string]$MessageId
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
        Write-Host ""
        Write-Host "Kies een actie voor deze e-mail:"
        Write-Host "1. Verwijder deze e-mail"
        Write-Host "2. Verplaats deze e-mail"
        Write-Host "3. Bekijk volledige body"
        Write-Host "4. Download bijlagen (indien aanwezig)"
        Write-Host "5. Terug naar zoekresultaten"

        $actionChoice = Read-Host "Kies een optie (1-5)"

        switch ($actionChoice) {
            "1" {
                $confirmDelete = Read-Host "Weet u zeker dat u deze e-mail permanent wilt verwijderen? (ja/nee)"
                if ($confirmDelete -eq 'ja') {
                    try {
                        Remove-MgUserMessage -UserId $UserId -MessageId $MessageId -ErrorAction Stop
                        Write-Host "E-mail succesvol verwijderd."
                        Write-Warning "De lokale cache (indien van toepassing) is mogelijk niet meer accuraat."
                    } catch {
                        Write-Error "Fout bij het verwijderen van de e-mail: $($_.Exception.Message)"
                    }
                } else {
                    Write-Host "Verwijderen geannuleerd."
                }
            }
            "2" {
                $destinationFolderId = Get-MailFolderSelection -UserId $UserId
                if ($destinationFolderId) {
                    $destinationFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $destinationFolderId -ErrorAction SilentlyContinue
                    $confirmMove = Read-Host "Weet u zeker dat u deze e-mail wilt verplaatsen naar '$($destinationFolder.DisplayName)'? (ja/nee)"
                    if ($confirmMove -eq 'ja') {
                        try {
                            Move-MgUserMessage -UserId $UserId -MessageId $MessageId -DestinationId $destinationFolderId -ErrorAction Stop
                            Write-Host "E-mail succesvol verplaatst naar '$($destinationFolder.DisplayName)'."
                            Write-Warning "De lokale cache (indien van toepassing) is mogelijk niet meer accuraat."
                        } catch {
                            Write-Error "Fout bij het verplaatsen van de e-mail: $($_.Exception.Message)"
                        }
                    } else {
                        Write-Host "Verplaatsen geannuleerd."
                    }
                } else {
                    Write-Host "Verplaatsen geannuleerd (geen doelmap geselecteerd)."
                }
            }
            "3" {
                Clear-Host
                Write-Host "Volledige body van e-mail (Onderwerp: $($message.Subject)):"
                Write-Host "----------------------------------------------------"
                if ($message.Body.ContentType -eq "html") {
                    Write-Warning "De body is in HTML-formaat. HTML-tags worden als platte tekst weergegeven."
                    # Voor een echte HTML-weergave zou een browser of een HTML-rendering component nodig zijn.
                    # Hier tonen we de ruwe HTML of de tekst-equivalent als die beschikbaar is.
                }
                Write-Host $message.Body.Content
                Write-Host "----------------------------------------------------"
                Read-Host "Druk op Enter om terug te keren naar het actiemenu"
                # Roep Show-EmailActionsMenu opnieuw aan om terug te keren naar het menu voor dezelfde e-mail
                Show-EmailActionsMenu -UserId $UserId -MessageId $MessageId
                return # Voorkom dubbele Read-Host aan het einde van de parent functie
            }
            "4" {
                if ($message.HasAttachments) {
                    Download-MessageAttachments -UserId $UserId -MessageId $MessageId
                } else {
                    Write-Host "Deze e-mail heeft geen bijlagen."
                }
                # Roep Show-EmailActionsMenu opnieuw aan om terug te keren naar het menu voor dezelfde e-mail
                Show-EmailActionsMenu -UserId $UserId -MessageId $MessageId
                return # Voorkom dubbele Read-Host aan het einde van de parent functie
            }
            "5" { 
                # Terugkeren gebeurt automatisch na de switch als er geen 'return' is in Search-Mail
                Write-Host "Terug naar zoekresultaten..."
                return 
            } 
            default { Write-Warning "Ongeldige keuze." }
        }

    } catch {
        Write-Error "Fout bij het ophalen of verwerken van e-mailacties: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) {
            Write-Error "StackTrace: $($_.ScriptStackTrace)"
        }
    }
    Read-Host "Druk op Enter om terug te keren naar het hoofdmenu (of vorige menu indien van toepassing)"
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
        [string]$MessageId
    )
    Clear-Host
    Write-Host "Bijlagen voor e-mail ID: $MessageId"
    Write-Host "-------------------------------------"
    
    try {
        $attachments = Get-MgUserMessageAttachment -UserId $UserId -MessageId $MessageId -ErrorAction Stop
        if ($null -eq $attachments -or $attachments.Count -eq 0) {
            Write-Warning "Geen bijlagen gevonden voor deze e-mail (ook al gaf HasAttachments 'true' aan)."
            Read-Host "Druk op Enter om terug te keren"
            return
        }

        $attachmentOptions = @{}
        $i = 1
        Write-Host "Beschikbare bijlagen:"
        foreach ($attachment in $attachments) {
            Write-Host "$i. $($attachment.Name) ($($attachment.Size) bytes, Type: $($attachment.ContentType))"
            $attachmentOptions[$i] = $attachment
            $i++
        }
        Write-Host "-------------------------------------"
        Write-Host "A. Download alle bijlagen"
        Write-Host "C. Annuleren"

        $choice = Read-Host "Kies een bijlage om te downloaden (nummer), A voor alles, of C om te annuleren"

        if ($choice -eq 'C' -or $choice -eq 'c') {
            Write-Host "Downloaden geannuleerd."
            return
        }

        $defaultDownloadPath = Join-Path -Path $PSScriptRoot -ChildPath "_downloads"
        $downloadPath = Read-Host "Voer het pad in voor de downloads (standaard: $defaultDownloadPath)"
        if ([string]::IsNullOrWhiteSpace($downloadPath)) {
            $downloadPath = $defaultDownloadPath
        }

        if (-not (Ensure-DownloadPath -Path $downloadPath)) {
            Read-Host "Druk op Enter om terug te keren"
            return
        }
        
        $attachmentsToDownload = New-Object System.Collections.Generic.List[object]
        if ($choice -eq 'A' -or $choice -eq 'a') {
            $attachments | ForEach-Object { $attachmentsToDownload.Add($_) }
        } elseif ($attachmentOptions.ContainsKey($choice)) {
            $attachmentsToDownload.Add($attachmentOptions[$choice])
        } else {
            Write-Warning "Ongeldige keuze."
            Read-Host "Druk op Enter om terug te keren"
            return
        }

        foreach ($attachment in $attachmentsToDownload) {
            $fileName = $attachment.Name
            # Sanitize filename (basic example, might need more robust sanitization)
            $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
            $regexInvalidChars = "[{0}]" -f ([regex]::Escape(-join $invalidChars))
            $safeFileName = $fileName -replace $regexInvalidChars, '_'
            
            $filePath = Join-Path -Path $downloadPath -ChildPath $safeFileName
            $counter = 1
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($safeFileName)
            $extension = [System.IO.Path]::GetExtension($safeFileName)
            
            while (Test-Path $filePath) {
                $newFileName = "{0}_{1}{2}" -f $baseName, $counter, $extension
                $filePath = Join-Path -Path $downloadPath -ChildPath $newFileName
                $counter++
            }

            Write-Host "Downloaden van '$($attachment.Name)' naar '$filePath'..."
            try {
                # Gebruik Invoke-MgGraphRequest om de raw content van de bijlage te krijgen
                $attachmentValueUri = "/users/$UserId/messages/$MessageId/attachments/$($attachment.Id)/`$value"
                $attachmentContent = Invoke-MgGraphRequest -Method GET -Uri $attachmentValueUri -ErrorAction Stop
                
                if ($attachmentContent) {
                    [System.IO.File]::WriteAllBytes($filePath, $attachmentContent)
                    Write-Host "Bijlage '$($attachment.Name)' succesvol opgeslagen als '$filePath'."
                } else {
                    Write-Warning "Invoke-MgGraphRequest gaf geen content terug voor bijlage '$($attachment.Name)'. Overslaan."
                }
            } catch {
                Write-Warning "Fout bij downloaden of opslaan van bijlage '$($attachment.Name)': $($_.Exception.Message)"
            }
        }

    } catch {
        Write-Error "Fout bij het ophalen of downloaden van bijlagen: $($_.Exception.Message)"
    }
    Read-Host "Druk op Enter om terug te keren"
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
    Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
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
        "1. Indexeer mailbox",
        "2. Overzicht van verzenders",
        "3. Beheer mails van specifieke afzender",
        "4. Zoek naar een mail",
        "5. Leeg 'Verwijderde Items'",
        "Q. Afsluiten"
    )
    $actionCodes = "1", "2", "3", "4", "5", "Q"
    
    $selectedItemIndex = 0
    $menuLoopActive = $true

    while ($menuLoopActive) {
        # Stel CGA-kleuren in voor het menu
        $Host.UI.RawUI.ForegroundColor = $cgaFgColor
        $Host.UI.RawUI.BackgroundColor = $cgaBgColor
        Clear-Host

        # Menu content
        $title = "OutlookBuddy - Hoofdmenu voor $UserEmail"
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
                "1" { Index-Mailbox -UserId $UserEmail }
                "2" { Show-SenderOverview -UserId $UserEmail }
                "3" { Manage-EmailsBySender -UserId $UserEmail }
                "4" { Search-Mail -UserId $UserEmail -IsTestRun:$TestRun.IsPresent }
                "5" { Empty-DeletedItemsFolder -UserId $UserEmail }
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
    }
    catch {
        throw "Kritiek: Fout tijdens het verbinden met Microsoft Graph: $($_.Exception.Message). Controleer de internetverbinding, de Microsoft Graph module installaties en de benodigde rechten/consent."
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
