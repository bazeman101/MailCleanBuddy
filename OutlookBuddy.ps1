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
    Ensure you have the necessary permissions (Microsoft Graph: Mail.Read) to access the specified mailbox.
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
    Write-Host "Menu Item 2: Overzicht van verzenders voor $UserId (Nog niet geïmplementeerd)"
    # TODO: Implement sender overview logic
    Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
}

function Manage-EmailsBySender {
    param($UserId)
    Write-Host "Menu Item 2.1-2.3: Beheer mails van specifieke afzender voor $UserId (Nog niet geïmplementeerd)"
    # TODO: Implement selection, deletion, and moving logic
    Read-Host "Druk op Enter om terug te keren naar het hoofdmenu"
}

function Search-Mail {
    param($UserId)
    Write-Host "Menu Item 3: Zoek naar een mail in $UserId (Nog niet geïmplementeerd)"
    # TODO: Implement mail search logic
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
    Write-Host "Q. Afsluiten"
    Write-Host "------------------------------------------"

    $choice = Read-Host "Kies een optie"

    switch ($choice) {
        "1" { Index-Mailbox -UserId $UserEmail }
        "2" { Show-SenderOverview -UserId $UserEmail }
        "3" { Manage-EmailsBySender -UserId $UserEmail }
        "4" { Search-Mail -UserId $UserEmail }
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
    $RequiredScopes = @("Mail.Read", "User.Read") # User.Read is often good to have for context

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
