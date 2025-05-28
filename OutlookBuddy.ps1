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
    [string]$MailboxEmail
)

# Script-level cache for sender information
$Script:SenderCache = $null 

# Placeholder functions for menu items
function Index-Mailbox {
    param($UserId)
    
    Write-Host "Starten met indexeren van mailbox voor $UserId..."
    $Script:SenderCache = @{} # Reset of initialiseer de cache

    try {
        Write-Host "Ophalen van berichten (dit kan even duren voor grote mailboxen)..."
        # Haal alleen de 'sender' eigenschap op voor efficiëntie
        # De -All parameter zorgt ervoor dat alle berichten worden opgehaald, ongeacht paginering
        $messages = Get-MgUserMessage -UserId $UserId -All -Property "sender" -ErrorAction Stop
        
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
                $senderAddress = $sender.Address.ToLowerInvariant() # Normaliseer e-mailadres
                $senderName = $sender.Name
                
                if ($Script:SenderCache.ContainsKey($senderAddress)) {
                    $Script:SenderCache[$senderAddress].Count++
                } else {
                    $Script:SenderCache[$senderAddress] = @{
                        Name  = $senderName
                        Count = 1
                    }
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
    
    Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
}

function Show-SenderOverview {
    param($UserId)
    Clear-Host
    Write-Host "Overzicht van verzenders voor $UserId"
    Write-Host "------------------------------------"

    if ($null -eq $Script:SenderCache -or $Script:SenderCache.Count -eq 0) {
        Write-Warning "De mailbox is nog niet geïndexeerd of de index is leeg."
        Write-Warning "Kies optie '1. Indexeer mailbox' in het hoofdmenu om de index op te bouwen."
        Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
        return
    }

    # Converteer de hashtable naar een array van custom objecten voor sortering en weergave
    $senderList = @()
    foreach ($key in $Script:SenderCache.Keys) {
        $senderList += [PSCustomObject]@{
            Email = $key
            Name  = $Script:SenderCache[$key].Name
            Count = $Script:SenderCache[$key].Count
        }
    }

    # Sorteer op aantal (aflopend) en dan op naam (oplopend)
    $sortedSenders = $senderList | Sort-Object -Property Count -Descending | Sort-Object -Property Name

    if ($sortedSenders.Count -eq 0) {
        Write-Host "Geen afzenders gevonden in de cache (dit zou niet moeten gebeuren als de indexering succesvol was)."
    } else {
        Write-Host "Afzenders gesorteerd op aantal e-mails (meeste eerst):"
        $sortedSenders | Format-Table -Property @{Name="Aantal"; Expression={$_.Count}; Width=7}, 
                                         @{Name="Naam"; Expression={$_.Name}; Width=40}, 
                                         @{Name="E-mailadres"; Expression={$_.Email}; Width=50} -AutoSize
    }
    
    Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
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
    Clear-Host
    Write-Host "Selecteer een doelmap:"
    Write-Host "-----------------------"
    try {
        # Haal alle mappen op, inclusief submappen (standaard gedrag van Get-MgUserMailFolder -All)
        # Sorteer op DisplayName voor consistentie
        $mailFolders = Get-MgUserMailFolder -UserId $UserId -All -ErrorAction Stop | Sort-Object DisplayName
        
        if ($null -eq $mailFolders -or $mailFolders.Count -eq 0) {
            Write-Warning "Geen mailmappen gevonden voor gebruiker $UserId."
            return $null
        }

        $folderOptions = @{}
        $i = 1
        Write-Host "Beschikbare mappen:"
        foreach ($folder in $mailFolders) {
            # Toon pad voor submappen voor duidelijkheid
            $displayPath = $folder.DisplayName
            $currentParentId = $folder.ParentFolderId
            $tempFolder = $folder # Om de originele folder niet te wijzigen
            while ($currentParentId) {
                $parentFolder = $mailFolders | Where-Object {$_.Id -eq $currentParentId} | Select-Object -First 1
                if ($parentFolder) {
                    $displayPath = "$($parentFolder.DisplayName) / $displayPath"
                    $currentParentId = $parentFolder.ParentFolderId
                } else {
                    $currentParentId = $null # Voorkom oneindige loop als ouder niet in de lijst staat
                }
            }

            Write-Host "$i. $displayPath (ID: $($folder.Id))"
            $folderOptions[$i] = $folder.Id
            $i++
        }
        Write-Host "-----------------------"
        Write-Host "C. Annuleren"

        while ($true) {
            $choice = Read-Host "Kies een mapnummer (of C om te annuleren)"
            if ($choice -eq 'C' -or $choice -eq 'c') {
                return $null
            }
            if ($folderOptions.ContainsKey($choice)) {
                return $folderOptions[$choice]
            } else {
                Write-Warning "Ongeldige keuze. Probeer opnieuw."
            }
        }
    } catch {
        Write-Error "Fout bij het ophalen van mailmappen: $($_.Exception.Message)"
        return $null
    }
}

function Search-Mail {
    param($UserId)
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
        $foundMessages = Get-MgUserMessage -UserId $UserId -Search $searchTerm -Top 100 -Property "subject,from,receivedDateTime,hasAttachments" -ErrorAction Stop
        
        if ($null -eq $foundMessages -or $foundMessages.Count -eq 0) {
            Write-Host "Geen e-mails gevonden die overeenkomen met de zoekterm '$searchTerm'."
        } else {
            Write-Host "$($foundMessages.Count) e-mail(s) gevonden:"
            Write-Host "----------------------------------------------------------------------------------------------------"
            $foundMessages | ForEach-Object {
                $fromAddress = if ($_.From -and $_.From.EmailAddress) { $_.From.EmailAddress.Address } else { "N/B" }
                $subject = if ($_.Subject) { $_.Subject } else { "(Geen onderwerp)" }
                $received = if ($_.ReceivedDateTime) { Get-Date $_.ReceivedDateTime -Format "yyyy-MM-dd HH:mm" } else { "N/B" }
                $attachments = if ($_.HasAttachments) { "Ja" } else { "Nee" }

                Write-Host ("Onderwerp    : {0}" -f $subject)
                Write-Host ("Van          : {0}" -f $fromAddress)
                Write-Host ("Ontvangen op : {0}" -f $received)
                Write-Host ("Bijlagen     : {0}" -f $attachments)
                Write-Host "ID           : $($_.Id)"
                Write-Host "----------------------------------------------------------------------------------------------------"
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

function Show-MainMenu {
    param (
        [string]$UserEmail
    )
    Clear-Host
    Write-Host "OutlookBuddy - Hoofdmenu voor $UserEmail"
    Write-Host "------------------------------------------"
    Write-Host "1. Indexeer mailbox"
    Write-Host "2. Overzicht van verzenders"
    Write-Host "3. Beheer mails van specifieke afzender"
    Write-Host "4. Zoek naar een mail"
    Write-Host "5. Leeg 'Verwijderde Items'"
    Write-Host "Q. Afsluiten"
    Write-Host "------------------------------------------"

    $choice = Read-Host "Kies een optie"

    switch ($choice) {
        "1" { Index-Mailbox -UserId $UserEmail }
        "2" { Show-SenderOverview -UserId $UserEmail }
        "3" { Manage-EmailsBySender -UserId $UserEmail }
        "4" { Search-Mail -UserId $UserEmail }
        "5" { Empty-DeletedItemsFolder -UserId $UserEmail }
        "Q" { Write-Host "Afsluiten..."; return $false } # Signal to exit loop
        default {
            Write-Warning "Ongeldige keuze. Probeer opnieuw."
            Read-Host "Druk op Enter om door te gaan"
        }
    }
    return $true # Signal to continue loop
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
