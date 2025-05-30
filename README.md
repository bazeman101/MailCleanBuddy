# MailCleanBuddy: Interactive Mailbox Manager for Microsoft 365

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/basw)

MailCleanBuddy is a PowerShell script that provides an interactive, menu-driven interface for managing your Microsoft 365 mailbox. It allows you to efficiently navigate your emails, perform bulk actions, and keep your mailbox tidy. This script was created partly to explore what's possible and partly as a genuine tool to help clean up mailboxes. It serves as an inspiration for others. If you use this script, a star or mention would be appreciated!

---

## English Documentation

### Purpose
The primary goal of this script is to quickly and easily clean up your Microsoft 365 mailbox via the command line. You can rapidly delete all emails from specific senders or domains. This script is partly an exploration of possibilities and partly a real aid for mailbox cleanup. It is intended to serve as inspiration for others.

### Main Features

*   **Multilingual Interface:** Supports Dutch (default), English, German, and French via a `localizations.json` file.
*   **Mailbox Indexing:** Creates a local cache of your mailbox for fast access to sender information and email metadata. This allows for quick overviews without repeatedly querying the server.
*   **Management by Sender/Domain:**
    *   View an overview of emails grouped by sender domain, sorted by count.
    *   Drill down into specific domains to view individual emails.
    *   Perform bulk actions (delete or move all emails) on all emails from a selected domain directly from the overview.
*   **Live Email Interactions:**
    *   Search for emails live based on keywords (searches subject, body, and sender).
    *   View the last 100 emails directly from the server.
*   **Email Actions:**
    *   Open and view email details, including a plain-text version of the body (converts HTML).
    *   Download attachments with a flexible and descriptive naming convention (`yyyy-MM-dd_senderdomain_senderextension_subject_attachmentname.ext`).
    *   Delete or move individual emails or multiple selected emails (using spacebar selection).
*   **Maintenance:** Empty the 'Deleted Items' folder.
*   **User-Friendly Menus:** Navigate easily with arrow keys, Enter, and hotkeys (e.g., 'Q' for Quit/Back, 'A' for Select All, 'N' for Deselect None, 'DEL' for delete, 'V' for move).
*   **Customizable Indexing:**
    *   `-TestRun` parameter to index only the latest 100 emails for quick testing.
    *   `-MaxEmailsToIndex` parameter to specify the exact number of latest emails to index.
*   **Console Customization:** Attempts to set an optimal console window size for better readability.

### Requirements

*   PowerShell (developed and tested on PowerShell 7+, but should be largely compatible with Windows PowerShell 5.1 with modern .NET).
*   Microsoft Graph PowerShell modules:
    *   `Microsoft.Graph.Authentication`
    *   `Microsoft.Graph.Mail`
    (The script will attempt to install these modules from the PowerShell Gallery if they are not found, using `Install-Module -Scope CurrentUser`).
*   Necessary Microsoft Graph permissions:
    *   `Mail.Read`: To read emails and their properties.
    *   `User.Read`: To get basic user information.
    *   `Mail.ReadWrite`: Required for deleting or moving emails, and emptying the deleted items folder.
    (You will be prompted by Microsoft to consent to these permissions on the first run or when scopes are missing).

### Usage

1.  **Clone or download the script** (`MailCleanBuddy.ps1`) and the `localizations.json` file into the same directory.
2.  **Open a PowerShell terminal** and navigate to the directory.
3.  **Run the script:**

    ```powershell
    .\MailCleanBuddy.ps1 -MailboxEmail "your-email@example.com"
    ```

    Replace `"your-email@example.com"` with the email address of the mailbox you want to manage.

### Parameters

*   `-MailboxEmail <string>`: (Mandatory) The email address of the mailbox to manage.
*   `-Language <string>`: (Optional) Specifies the UI language. Supported: `nl` (Dutch - default), `en` (English), `de` (German), `fr` (French). Example: `-Language en`.
*   `-TestRun <switch>`: (Optional) If specified, the script will only index the latest 100 emails. This is useful for quick testing.
*   `-MaxEmailsToIndex <int>`: (Optional) Specifies the maximum number of newest emails to index. If this value is greater than 0, it overrides the `-TestRun` switch for the number of emails to index. Default is 0 (uses `-TestRun` logic or full indexing).

### Key Functionalities (Overview)

*   **`Load-LocalizationStrings` / `Get-LocStr`**: Handles loading and retrieving translated UI strings from `localizations.json`.
*   **`Get-CacheFilePath` / `Load-LocalCache` / `Save-LocalCache`**: Manage the local JSON cache for sender/domain information.
*   **`Index-Mailbox`**: Fetches emails from the server (all, or limited by `-TestRun`/`-MaxEmailsToIndex`), processes them to group by sender domain, and populates the cache. Uses MAPI properties for reliable size and attachment detection.
*   **`Show-MainMenu`**: Displays the main interactive menu with options to navigate to different functionalities.
*   **`Show-SenderOverview`**: Displays a list of sender domains from the cache, sorted by email count. Allows opening a domain's emails or performing bulk actions (delete/move all from domain).
*   **`Show-EmailsFromSelectedSender`**: Called from `Show-SenderOverview`, it prepares and displays emails for a specific domain using `Show-StandardizedEmailListView`.
*   **`Show-StandardizedEmailListView`**: A generic function to display lists of emails (from cache, search results, or recent emails). Handles navigation (scrolling, selection with spacebar, select all/none), and invoking actions (Enter to open, DEL to delete, V to move).
*   **`Perform-ActionOnMultipleEmails`**: Handles the menu and logic for deleting or moving multiple selected emails.
*   **`Perform-ActionOnAllSenderEmails`**: Handles the menu and logic for deleting or moving all emails from a specific sender domain (from cache).
*   **`Show-EmailActionsMenu`**: Displays an action menu for a single opened email (fetched live from the server). Allows deleting, moving, viewing the body, or downloading attachments for that specific email. Updates cache if necessary.
*   **`Show-EmailBody`**: Fetches and displays the full body content of an email, converting HTML to plain text.
*   **`Download-MessageAttachments`**: Lists attachments for an email and allows downloading selected or all attachments with a descriptive naming convention.
*   **`Search-Mail`**: Prompts for a search term and displays matching emails using `Show-StandardizedEmailListView`.
*   **`Show-RecentEmails`**: Fetches and displays the last 100 emails using `Show-StandardizedEmailListView`.
*   **`Empty-DeletedItemsFolder`**: Empties the "Deleted Items" folder after confirmation.
*   **`Get-MailFolderSelection`**: Provides an interactive menu to select a mail folder (used for moving emails).
*   **`Get-Confirmation`**: A reusable function to display a Yes/No confirmation prompt.
*   **`Convert-HtmlToPlainText`**: A helper function to strip HTML tags and convert HTML content to a more readable plain text format.
*   **Module & Graph Connection Handling**: The script checks for required Microsoft Graph modules, attempts installation if missing, and manages the connection to Microsoft Graph, including requesting necessary scopes.

### Localization
The script supports multiple languages for its user interface. Translations are stored in the `localizations.json` file.
Currently supported:
*   Dutch (`nl` - default)
*   English (`en`)
*   German (`de`)
*   French (`fr`)

You can select a language using the `-Language` parameter, e.g., `.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -Language en`.
If a translation key is missing for a selected language, it will fall back to the key itself. If a language is not found, it falls back to Dutch.

### Contributing / Feedback
This script was developed as a personal project and for learning purposes. If you find it useful, have suggestions, or find bugs, feel free to:
*   Star the repository!
*   Open an issue for bugs or feature requests.
*   Submit a pull request with improvements.

---

## Nederlandse Documentatie

### Doel
Het primaire doel van dit script is om snel en eenvoudig je Microsoft 365 mailbox op te ruimen via de commando regel. Je kunt hiermee zeer snel alle e-mails van bepaalde afzenders of domeinen verwijderen. Dit script is deels gemaakt om te kijken wat er mogelijk is en deels als een daadwerkelijke hulp bij het opruimen van mailboxen. Het dient als inspiratie voor anderen.

### Belangrijkste Functionaliteiten

*   **Meertalige Interface:** Ondersteunt Nederlands (standaard), Engels, Duits en Frans via een `localizations.json` bestand.
*   **Mailbox Indexering:** Creëert een lokale cache van je mailbox voor snelle toegang tot afzenderinformatie en e-mail metadata. Dit maakt snelle overzichten mogelijk zonder herhaaldelijk de server te bevragen.
*   **Beheer per Afzender/Domein:**
    *   Bekijk een overzicht van e-mails gegroepeerd per afzenderdomein, gesorteerd op aantal.
    *   Zoom in op specifieke domeinen om individuele e-mails te bekijken.
    *   Voer bulkacties uit (verwijder of verplaats alle e-mails) op alle e-mails van een geselecteerd domein direct vanuit het overzicht.
*   **Live E-mail Interacties:**
    *   Zoek live naar e-mails op basis van trefwoorden (zoekt in onderwerp, body en afzender).
    *   Bekijk de laatste 100 e-mails direct van de server.
*   **E-mail Acties:**
    *   Open en bekijk e-maildetails, inclusief een platte-tekstversie van de body (converteert HTML).
    *   Download bijlagen met een flexibele en beschrijvende naamgevingsconventie (`jjjj-MM-dd_afzenderdomein_afzenderextensie_onderwerp_bijlagenaam.ext`).
    *   Verwijder of verplaats individuele e-mails of meerdere geselecteerde e-mails (middels spatiebalkselectie).
*   **Onderhoud:** Leeg de map 'Verwijderde Items'.
*   **Gebruiksvriendelijke Menu's:** Navigeer eenvoudig met pijltjestoetsen, Enter, en sneltoetsen (bijv. 'Q' voor Afsluiten/Terug, 'A' voor Alles Selecteren, 'N' voor Selectie Opheffen, 'DEL' voor verwijderen, 'V' voor verplaatsen).
*   **Aanpasbare Indexering:**
    *   `-TestRun` parameter om alleen de laatste 100 e-mails te indexeren voor snel testen.
    *   `-MaxEmailsToIndex` parameter om het exacte aantal nieuwste e-mails te specificeren voor indexering.
*   **Console Aanpassing:** Probeert een optimale console venstergrootte in te stellen voor betere leesbaarheid.

### Vereisten

*   PowerShell (ontwikkeld en getest op PowerShell 7+, maar zou grotendeels compatibel moeten zijn met Windows PowerShell 5.1 met modern .NET).
*   Microsoft Graph PowerShell modules:
    *   `Microsoft.Graph.Authentication`
    *   `Microsoft.Graph.Mail`
    (Het script probeert deze modules te installeren vanuit de PowerShell Gallery indien ze niet gevonden worden, middels `Install-Module -Scope CurrentUser`).
*   Benodigde Microsoft Graph permissies:
    *   `Mail.Read`: Om e-mails en hun eigenschappen te lezen.
    *   `User.Read`: Om basis gebruikersinformatie op te halen.
    *   `Mail.ReadWrite`: Vereist voor het verwijderen of verplaatsen van e-mails, en het legen van de map verwijderde items.
    (Je zult door Microsoft gevraagd worden om toestemming te geven voor deze permissies bij de eerste uitvoering of wanneer scopes ontbreken).

### Gebruik

1.  **Kloon of download het script** (`MailCleanBuddy.ps1`) en het `localizations.json` bestand naar dezelfde map.
2.  **Open een PowerShell terminal** en navigeer naar de map.
3.  **Voer het script uit:**

    ```powershell
    .\MailCleanBuddy.ps1 -MailboxEmail "jouw-email@example.com"
    ```

    Vervang `"jouw-email@example.com"` met het e-mailadres van de mailbox die je wilt beheren.

### Parameters

*   `-MailboxEmail <string>`: (Verplicht) Het e-mailadres van de mailbox die beheerd moet worden.
*   `-Language <string>`: (Optioneel) Specificeert de UI-taal. Ondersteund: `nl` (Nederlands - standaard), `en` (Engels), `de` (Duits), `fr` (Frans). Voorbeeld: `-Language en`.
*   `-TestRun <switch>`: (Optioneel) Indien gespecificeerd, indexeert het script alleen de laatste 100 e-mails. Handig voor snel testen.
*   `-MaxEmailsToIndex <int>`: (Optioneel) Specificeert het maximale aantal nieuwste e-mails dat geïndexeerd moet worden. Als deze waarde groter is dan 0, overschrijft dit de `-TestRun` switch voor het aantal te indexeren e-mails. Standaard is 0 (gebruikt `-TestRun` logica of volledige indexering).

### Kernfunctionaliteiten (Overzicht)

*   **`Load-LocalizationStrings` / `Get-LocStr`**: Behandelt het laden en ophalen van vertaalde UI-strings uit `localizations.json`.
*   **`Get-CacheFilePath` / `Load-LocalCache` / `Save-LocalCache`**: Beheren de lokale JSON-cache voor afzender-/domeininformatie.
*   **`Index-Mailbox`**: Haalt e-mails op van de server (alle, of beperkt door `-TestRun`/`-MaxEmailsToIndex`), verwerkt ze om te groeperen per afzenderdomein, en vult de cache. Gebruikt MAPI-eigenschappen voor betrouwbare detectie van grootte en bijlagen.
*   **`Show-MainMenu`**: Toont het interactieve hoofdmenu met opties om naar verschillende functionaliteiten te navigeren.
*   **`Show-SenderOverview`**: Toont een lijst van afzenderdomeinen uit de cache, gesorteerd op e-mailaantal. Maakt het mogelijk om e-mails van een domein te openen of bulkacties uit te voeren (verwijder/verplaats alles van domein).
*   **`Show-EmailsFromSelectedSender`**: Aangeroepen vanuit `Show-SenderOverview`, bereidt e-mails voor een specifiek domein voor en toont deze middels `Show-StandardizedEmailListView`.
*   **`Show-StandardizedEmailListView`**: Een generieke functie om lijsten van e-mails te tonen (uit cache, zoekresultaten, of recente e-mails). Behandelt navigatie (scrollen, selectie met spatiebalk, alles selecteren/deselecteren), en het aanroepen van acties (Enter om te openen, DEL om te verwijderen, V om te verplaatsen).
*   **`Perform-ActionOnMultipleEmails`**: Behandelt het menu en de logica voor het verwijderen of verplaatsen van meerdere geselecteerde e-mails.
*   **`Perform-ActionOnAllSenderEmails`**: Behandelt het menu en de logica voor het verwijderen of verplaatsen van alle e-mails van een specifiek afzenderdomein (uit cache).
*   **`Show-EmailActionsMenu`**: Toont een actiemenu voor een enkele geopende e-mail (live opgehaald van de server). Maakt het mogelijk om te verwijderen, verplaatsen, de body te bekijken, of bijlagen te downloaden voor die specifieke e-mail. Werkt de cache bij indien nodig.
*   **`Show-EmailBody`**: Haalt de volledige body-inhoud van een e-mail op en toont deze, converteert HTML naar platte tekst.
*   **`Download-MessageAttachments`**: Toont bijlagen voor een e-mail en maakt het mogelijk geselecteerde of alle bijlagen te downloaden met een beschrijvende naamgevingsconventie.
*   **`Search-Mail`**: Vraagt om een zoekterm en toont overeenkomende e-mails middels `Show-StandardizedEmailListView`.
*   **`Show-RecentEmails`**: Haalt de laatste 100 e-mails op en toont deze middels `Show-StandardizedEmailListView`.
*   **`Empty-DeletedItemsFolder`**: Leegt de map 'Verwijderde Items' na bevestiging.
*   **`Get-MailFolderSelection`**: Biedt een interactief menu om een e-mailmap te selecteren (gebruikt voor het verplaatsen van e-mails).
*   **`Get-Confirmation`**: Een herbruikbare functie om een Ja/Nee bevestigingsprompt te tonen.
*   **`Convert-HtmlToPlainText`**: Een hulpfunctie om HTML-tags te verwijderen en HTML-inhoud om te zetten naar een beter leesbaar platte-tekstformaat.
*   **Module & Graph Connectie Beheer**: Het script controleert op vereiste Microsoft Graph modules, probeert installatie indien nodig, en beheert de verbinding met Microsoft Graph, inclusief het aanvragen van de benodigde scopes.

### Lokalisatie
Het script ondersteunt meerdere talen voor de gebruikersinterface. Vertalingen worden opgeslagen in het `localizations.json` bestand.
Momenteel ondersteund:
*   Nederlands (`nl` - standaard)
*   Engels (`en`)
*   Duits (`de`)
*   Frans (`fr`)

Je kunt een taal selecteren met de `-Language` parameter, bijv., `.\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -Language en`.
Als een vertaalsleutel ontbreekt voor een geselecteerde taal, valt het script terug op de sleutel zelf. Als een taal niet wordt gevonden, valt het terug op Nederlands.

### Bijdragen / Feedback
Dit script is ontwikkeld als een persoonlijk project en voor leerdoeleinden. Als je het nuttig vindt, suggesties hebt, of bugs vindt, voel je vrij om:
*   De repository een ster te geven!
*   Een issue te openen voor bugs of feature requests.
*   Een pull request in te dienen met verbeteringen.

---
[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/basw)
