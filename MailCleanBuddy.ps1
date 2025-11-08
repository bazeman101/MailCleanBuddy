<#
.SYNOPSIS
    MailCleanBuddy - Interactive Mailbox Manager for Microsoft 365 (Modular Version)
.DESCRIPTION
    This script provides an interactive, menu-driven interface for managing your Microsoft 365 mailbox.
    It allows you to efficiently navigate your emails, perform bulk actions, and keep your mailbox tidy.

    This is the modularized version with enhanced features:
    - Bulk attachment download with time filters
    - Email export to EML/MSG format
    - Move senders to subfolders
    - Improved attachment download (FIXED base64 decoding)

.PARAMETER MailboxEmail
    The email address of the mailbox to manage.
.PARAMETER Language
    UI language (nl, en, de, fr). Default: nl
.PARAMETER TestRun
    Index only the latest 100 emails for quick testing.
.PARAMETER MaxEmailsToIndex
    Specify the maximum number of newest emails to index. Default: 0 (full indexing)

.EXAMPLE
    .\MailCleanBuddy.ps1 -MailboxEmail "user@example.com"

.EXAMPLE
    .\MailCleanBuddy.ps1 -MailboxEmail "user@example.com" -Language en -TestRun

.NOTES
    Version: 2.3 (Full Feature Suite)
    Requires: PowerShell 7+ (compatible with Windows PowerShell 5.1)
    Modules: Microsoft.Graph.Authentication, Microsoft.Graph.Mail
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "The email address of the mailbox to manage.")]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern('^[\w\-\.]+@([\w\-]+\.)+[\w\-]{2,}$', ErrorMessage = "Invalid email address format")]
    [string]$MailboxEmail,

    [Parameter(Mandatory = $false)]
    [switch]$TestRun,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 10000)]
    [int]$MaxEmailsToIndex = 0,

    [Parameter(Mandatory = $false, HelpMessage = "UI language (nl, en, de, fr). Default: nl")]
    [ValidateSet('nl', 'en', 'de', 'fr')]
    [string]$Language = "nl"
)

#region Module Imports
Write-Host "Loading MailCleanBuddy modules..." -ForegroundColor Cyan

$ModulesPath = Join-Path $PSScriptRoot "Modules"

# Import all modules
try {
    Import-Module (Join-Path $ModulesPath "Utilities\Localization.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Utilities\Helpers.psm1") -Force
    Import-Module (Join-Path $ModulesPath "UI\ColorScheme.psm1") -Force
    Import-Module (Join-Path $ModulesPath "UI\Display.psm1") -Force
    Import-Module (Join-Path $ModulesPath "UI\MenuSystem.psm1") -Force
    Import-Module (Join-Path $ModulesPath "UI\EmailListView.psm1") -Force
    Import-Module (Join-Path $ModulesPath "UI\EmailViewer.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Core\GraphApiService.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Core\CacheManager.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\EmailActions.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\EmailSearch.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\AttachmentManager.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\MessageExport.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\UnsubscribeManager.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\DuplicateDetector.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\LargeAttachmentManager.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\EmailArchiver.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\ThreadAnalyzer.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\SmartOrganizer.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Analytics\AnalyticsDashboard.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Analytics\AttachmentStats.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Utilities\VIPManager.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Utilities\HeaderAnalyzer.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Utilities\HealthMonitor.psm1") -Force
    Import-Module (Join-Path $ModulesPath "EmailOperations\AdvancedSearch.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Security\ThreatDetector.psm1") -Force
    Import-Module (Join-Path $ModulesPath "Integration\CalendarSync.psm1") -Force

    Write-Host "Modules loaded successfully!" -ForegroundColor Green
} catch {
    Write-Error "Failed to load modules: $($_.Exception.Message)"
    Write-Error "Please ensure all module files are present in the Modules directory."
    exit 1
}
#endregion

#region Initialization
# Initialize localization
Initialize-Localization -SelectedLang $Language -FilePath (Join-Path $PSScriptRoot "localizations.json")

# Set console size
Set-ConsoleSize -Width 150 -Height 55

# Set default colors
Set-DefaultColors

Write-Host ""
Write-Host "=== MailCleanBuddy v3.0 (Complete Feature Suite) ===" -ForegroundColor Green
Write-Host "Mailbox: $MailboxEmail" -ForegroundColor Cyan
Write-Host ""
#endregion

#region Graph Module Check and Installation
$moduleCheck = Test-GraphModules
if (-not $moduleCheck.AllInstalled) {
    Write-Host "Missing Microsoft Graph modules detected." -ForegroundColor Yellow
    Write-Host "The following modules will be installed: $($moduleCheck.MissingModules -join ', ')" -ForegroundColor Yellow

    try {
        Install-GraphModules -Modules $moduleCheck.MissingModules
    } catch {
        Write-Error "Critical: Could not install required modules. Please install manually with:"
        Write-Error "Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Mail -Scope CurrentUser"
        exit 1
    }
}
#endregion

#region Graph Connection
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
$connected = Connect-GraphService -Scopes @("Mail.Read", "Mail.ReadWrite", "User.Read")

if (-not $connected) {
    Write-Error "Failed to connect to Microsoft Graph. Exiting."
    exit 1
}
#endregion

#region Cache Initialization
# Set cache file path
$cacheFilePath = Get-CacheFilePath -MailboxEmail $MailboxEmail -BasePath $PSScriptRoot

# Try to load existing cache
$cacheLoaded = Import-LocalCache -FilePath $cacheFilePath

if ($cacheLoaded) {
    Write-Host "Local cache loaded successfully. Skipping full mailbox indexing for faster startup." -ForegroundColor Green
    Write-Host "Use menu option 'R' to refresh the index from the server if needed." -ForegroundColor Yellow
} else {
    Write-Host "No valid cache found. Starting automatic mailbox indexing..." -ForegroundColor Cyan

    $indexParams = @{
        UserId = $MailboxEmail
    }

    if ($MaxEmailsToIndex -gt 0) {
        $indexParams.MaxEmailsToIndex = $MaxEmailsToIndex
    } elseif ($TestRun) {
        $indexParams.TestMode = $true
    }

    Build-MailboxIndex @indexParams

    Write-Host "Automatic indexing completed." -ForegroundColor Green
}
#endregion

#region Main Menu Functions

function Show-MainMenuOptions {
    param([string]$UserEmail)

    Clear-Host
    Show-Header -Title (Get-LocalizedString "mainMenu_title" -FormatArgs $UserEmail) -Width 100

    Write-Host ""
    Write-Host "  $(Get-LocalizedString 'mainMenu_option1')" -ForegroundColor Green
    Write-Host "  $(Get-LocalizedString 'mainMenu_option2')" -ForegroundColor Green
    Write-Host "  $(Get-LocalizedString 'mainMenu_option3')" -ForegroundColor Green
    Write-Host "  $(Get-LocalizedString 'mainMenu_option4')" -ForegroundColor Green
    Write-Host "  $(Get-LocalizedString 'mainMenu_option5')" -ForegroundColor Green
    Write-Host ""
    Write-Host "  A. $(Get-LocalizedString 'mainMenu_newOption1')" -ForegroundColor Magenta
    Write-Host "  B. $(Get-LocalizedString 'mainMenu_newOption2')" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  $(Get-LocalizedString 'mainMenu_optionR')" -ForegroundColor Yellow
    Write-Host "  $(Get-LocalizedString 'mainMenu_optionQ')" -ForegroundColor Red
    Write-Host ""

    $selection = Read-Host "Select an option"
    return $selection
}

function Show-SmartFeaturesMenu {
    param([string]$UserEmail)

    $continue = $true
    while ($continue) {
        Clear-Host
        Show-Header -Title (Get-LocalizedString "smartMenu_title") -Width 100

        Write-Host ""
        Write-Host "  $(Get-LocalizedString 'smartMenu_option1')" -ForegroundColor Green
        Write-Host "  $(Get-LocalizedString 'smartMenu_option2')" -ForegroundColor Green
        Write-Host "  $(Get-LocalizedString 'smartMenu_option3')" -ForegroundColor Green
        Write-Host "  4. Large Attachment Manager" -ForegroundColor Green
        Write-Host "  5. Email Archiver (Retentie Beleid)" -ForegroundColor Green
        Write-Host "  6. Thread Analyzer" -ForegroundColor Green
        Write-Host "  7. Smart Folder Organizer" -ForegroundColor Green
        Write-Host "  8. VIP Sender Manager" -ForegroundColor Green
        Write-Host "  9. Email Header Analyzer" -ForegroundColor Green
        Write-Host "  10. Advanced Email Search" -ForegroundColor Cyan
        Write-Host "  11. Mailbox Health Monitor" -ForegroundColor Cyan
        Write-Host "  12. Threat Detection (Security)" -ForegroundColor Cyan
        Write-Host "  13. Calendar Integration" -ForegroundColor Cyan
        Write-Host "  14. Attachment Statistics" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  $(Get-LocalizedString 'smartMenu_optionBack')" -ForegroundColor Yellow
        Write-Host ""

        $selection = Read-Host "Select an option"

        switch ($selection.ToUpper()) {
            "1" {
                Show-AnalyticsDashboard -UserEmail $UserEmail
            }
            "2" {
                Show-UnsubscribeOpportunities -UserEmail $UserEmail
            }
            "3" {
                Show-DuplicateEmailsManager -UserEmail $UserEmail
            }
            "4" {
                Show-LargeAttachmentManager -UserEmail $UserEmail
            }
            "5" {
                Show-EmailArchiver -UserEmail $UserEmail
            }
            "6" {
                Show-ThreadAnalyzer -UserEmail $UserEmail
            }
            "7" {
                Show-SmartOrganizer -UserEmail $UserEmail
            }
            "8" {
                Show-VIPManager -UserEmail $UserEmail
            }
            "9" {
                Show-HeaderAnalyzer -UserEmail $UserEmail
            }
            "10" {
                Show-AdvancedSearch -UserEmail $UserEmail
            }
            "11" {
                Show-HealthMonitor -UserEmail $UserEmail
            }
            "12" {
                Show-ThreatDetector -UserEmail $UserEmail
            }
            "13" {
                Show-CalendarSync -UserEmail $UserEmail
            }
            "14" {
                Show-AttachmentStats -UserEmail $UserEmail
            }
            "B" {
                $continue = $false
            }
            default {
                Write-Host (Get-LocalizedString "mainMenu_actionUnknown" -FormatArgs $selection) -ForegroundColor Red
                Wait-EnterKey
            }
        }
    }
}

function Show-BulkOperationsMenu {
    param([string]$UserEmail)

    $continue = $true
    while ($continue) {
        Clear-Host
        Show-Header -Title (Get-LocalizedString "bulkMenu_title") -Width 100

        Write-Host ""
        Write-Host (Get-LocalizedString "bulkMenu_option1") -ForegroundColor Green
        Write-Host (Get-LocalizedString "bulkMenu_option2") -ForegroundColor Green
        Write-Host (Get-LocalizedString "bulkMenu_option3") -ForegroundColor Green
        Write-Host ""
        Write-Host (Get-LocalizedString "bulkMenu_optionBack") -ForegroundColor Yellow
        Write-Host ""

        $selection = Read-Host "Select an option"

        switch ($selection.ToUpper()) {
            "1" {
                Invoke-BulkAttachmentDownload -UserEmail $UserEmail
            }
            "2" {
                Invoke-EmailExport -UserEmail $UserEmail
            }
            "3" {
                Invoke-MoveToSubfolder -UserEmail $UserEmail
            }
            "B" {
                $continue = $false
            }
            default {
                Write-Host (Get-LocalizedString "mainMenu_actionUnknown" -FormatArgs $selection) -ForegroundColor Red
                Wait-EnterKey
            }
        }
    }
}

function Show-EmailManagementMenu {
    param([string]$UserEmail)

    $continue = $true
    while ($continue) {
        Clear-Host
        Show-Header -Title (Get-LocalizedString "emailMenu_title") -Width 100

        Write-Host ""
        Write-Host (Get-LocalizedString "emailMenu_option1") -ForegroundColor Green
        Write-Host (Get-LocalizedString "emailMenu_option2") -ForegroundColor Green
        Write-Host (Get-LocalizedString "emailMenu_option3") -ForegroundColor Green
        Write-Host (Get-LocalizedString "emailMenu_option4") -ForegroundColor Green
        Write-Host (Get-LocalizedString "emailMenu_option5") -ForegroundColor Green
        Write-Host ""
        Write-Host (Get-LocalizedString "emailMenu_optionBack") -ForegroundColor Yellow
        Write-Host ""

        $selection = Read-Host "Select an option"

        switch ($selection.ToUpper()) {
            "1" {
                Show-SenderOverviewMenu -UserEmail $UserEmail
            }
            "2" {
                Show-SenderEmailsMenu -UserEmail $UserEmail
            }
            "3" {
                Invoke-EmailSearch -UserEmail $UserEmail
            }
            "4" {
                Show-RecentEmails -UserEmail $UserEmail -Count 100
            }
            "5" {
                $confirm = Show-Confirmation -Message "Are you sure you want to empty the Deleted Items folder? This cannot be undone."
                if ($confirm) {
                    Write-Host "Emptying Deleted Items folder..." -ForegroundColor Cyan
                    $result = Clear-GraphDeletedItems -UserId $UserEmail
                    if ($result) {
                        Write-Host "Deleted Items folder emptied successfully." -ForegroundColor Green
                    } else {
                        Write-Host "Failed to empty Deleted Items folder. See error above." -ForegroundColor Red
                    }
                } else {
                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                }
                Wait-EnterKey
            }
            "B" {
                $continue = $false
            }
            default {
                Write-Host (Get-LocalizedString "mainMenu_actionUnknown" -FormatArgs $selection) -ForegroundColor Red
                Wait-EnterKey
            }
        }
    }
}

function Invoke-BulkAttachmentDownload {
    param([string]$UserEmail)

    Clear-Host
    Show-Header -Title (Get-LocalizedString "bulkAttachments_title") -Width 100

    # Select time filter
    $timeFilter = Show-TimeFilterMenu -Title (Get-LocalizedString "bulkAttachments_selectTimeFilter")
    if (-not $timeFilter) {
        Write-Host (Get-LocalizedString "moveToSubfolder_cancelled") -ForegroundColor Yellow
        Wait-EnterKey
        return
    }

    # Get save path
    $defaultPath = Join-Path $PSScriptRoot "_attachments"
    $savePath = Read-Host (Get-LocalizedString "bulkAttachments_enterPath" -FormatArgs $defaultPath)
    if ([string]::IsNullOrWhiteSpace($savePath)) {
        $savePath = $defaultPath
    }

    # Optional: sender domain filter
    $senderDomain = Read-Host (Get-LocalizedString "bulkAttachments_enterDomain")

    # Optional: file types filter
    $fileTypesInput = Read-Host (Get-LocalizedString "bulkAttachments_enterFileTypes")
    $fileTypes = if ([string]::IsNullOrWhiteSpace($fileTypesInput)) { $null } else { $fileTypesInput -split ',' | ForEach-Object { $_.Trim() } }

    # Skip duplicates?
    $skipDupsInput = Read-Host (Get-LocalizedString "bulkAttachments_skipDuplicates")
    $skipDuplicates = $skipDupsInput -match '^(y|yes|j|ja)$'

    Write-Host ""
    Write-Host (Get-LocalizedString "bulkAttachments_starting") -ForegroundColor Cyan

    $params = @{
        UserId = $UserEmail
        TimeFilter = $timeFilter
        SavePath = $savePath
    }
    if (-not [string]::IsNullOrWhiteSpace($senderDomain)) { $params.SenderDomain = $senderDomain }
    if ($fileTypes) { $params.FileTypes = $fileTypes }
    if ($skipDuplicates) { $params.SkipDuplicates = $true }

    Get-BulkAttachments @params

    Write-Host ""
    Write-Host (Get-LocalizedString "bulkAttachments_complete") -ForegroundColor Green
    Wait-EnterKey
}

function Invoke-EmailExport {
    param([string]$UserEmail)

    Clear-Host
    Show-Header -Title (Get-LocalizedString "exportEmails_title") -Width 100

    # Select format
    Write-Host (Get-LocalizedString "exportEmails_selectFormat") -ForegroundColor Yellow
    $formatChoice = Read-Host "Choice"
    $format = if ($formatChoice -eq "2") { "MSG" } else { "EML" }

    # Select time filter
    $timeFilter = Show-TimeFilterMenu -Title (Get-LocalizedString "exportEmails_selectTimeFilter")
    if (-not $timeFilter) {
        Write-Host (Get-LocalizedString "moveToSubfolder_cancelled") -ForegroundColor Yellow
        Wait-EnterKey
        return
    }

    # Get save path
    $defaultPath = Join-Path $PSScriptRoot "_exported_emails"
    $savePath = Read-Host (Get-LocalizedString "exportEmails_enterPath" -FormatArgs $defaultPath)
    if ([string]::IsNullOrWhiteSpace($savePath)) {
        $savePath = $defaultPath
    }

    # Optional: sender domain filter
    $senderDomain = Read-Host (Get-LocalizedString "exportEmails_enterDomain")

    Write-Host ""
    Write-Host (Get-LocalizedString "exportEmails_starting") -ForegroundColor Cyan

    $params = @{
        UserId = $UserEmail
        TimeFilter = $timeFilter
        SavePath = $savePath
        Format = $format
    }
    if (-not [string]::IsNullOrWhiteSpace($senderDomain)) { $params.SenderDomain = $senderDomain }

    Export-BulkEmails @params

    Write-Host ""
    Write-Host (Get-LocalizedString "exportEmails_complete") -ForegroundColor Green
    Wait-EnterKey
}

function Invoke-MoveToSubfolder {
    param([string]$UserEmail)

    Clear-Host
    Show-Header -Title (Get-LocalizedString "moveToSubfolder_title") -Width 100

    # Get sender domain
    $senderDomain = Read-Host (Get-LocalizedString "moveToSubfolder_enterDomain")
    if ([string]::IsNullOrWhiteSpace($senderDomain)) {
        Write-Host (Get-LocalizedString "moveToSubfolder_cancelled") -ForegroundColor Yellow
        Wait-EnterKey
        return
    }

    # Get subfolder name
    $subfolderName = Read-Host (Get-LocalizedString "moveToSubfolder_enterSubfolderName")
    if ([string]::IsNullOrWhiteSpace($subfolderName)) {
        Write-Host (Get-LocalizedString "moveToSubfolder_cancelled") -ForegroundColor Yellow
        Wait-EnterKey
        return
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "moveToSubfolder_starting") -ForegroundColor Cyan

    # First run in preview mode
    Move-SenderToSubfolder -UserId $UserEmail -SenderDomain $senderDomain -SubfolderName $subfolderName -PreviewOnly

    Write-Host ""
    $confirm = Read-Host (Get-LocalizedString "moveToSubfolder_confirmMove")
    if ($confirm -match '^(y|yes|j|ja)$') {
        # Actually move
        Move-SenderToSubfolder -UserId $UserEmail -SenderDomain $senderDomain -SubfolderName $subfolderName

        Write-Host ""
        Write-Host (Get-LocalizedString "moveToSubfolder_complete") -ForegroundColor Green
    } else {
        Write-Host (Get-LocalizedString "moveToSubfolder_cancelled") -ForegroundColor Yellow
    }

    Wait-EnterKey
}

function Show-SenderOverviewMenu {
    param([string]$UserEmail)

    $cache = Get-SenderCache

    if ($null -eq $cache -or $cache.Count -eq 0) {
        Show-WarningMessage (Get-LocalizedString "senderOverview_notIndexedOrEmpty")
        Wait-EnterKey
        return
    }

    # Convert cache to array of objects
    $senderList = @()
    foreach ($domainKey in $cache.Keys) {
        $senderList += [PSCustomObject]@{
            Domain = $domainKey
            Count = $cache[$domainKey].Count
            DisplayText = "$($cache[$domainKey].Count.ToString().PadLeft(6)) | $domainKey"
        }
    }

    # Sort by count descending
    $senderList = $senderList | Sort-Object Count -Descending

    # Show selectable list
    $selected = Show-SelectableList -Title (Get-LocalizedString "senderOverview_title") `
                                    -Items $senderList `
                                    -DisplayProperty "DisplayText" `
                                    -PageSize 30

    if ($selected) {
        # User selected a domain - show emails from that domain
        $senderData = $cache[$selected.Domain]

        if ($senderData -and $senderData.Messages) {
            Write-Host "`nLoading emails from: $($selected.Domain)" -ForegroundColor Cyan
            Start-Sleep -Seconds 1

            # Prepare messages for display
            $messagesForView = @()
            foreach ($msg in $senderData.Messages) {
                # Cache messages use MessageId property, not Id
                # Handle both hashtables (from cache) and PSCustomObjects
                $msgId = $null
                if ($msg.MessageId) {
                    $msgId = $msg.MessageId
                } elseif ($msg.Id) {
                    $msgId = $msg.Id
                }

                $messagesForView += [PSCustomObject]@{
                    Id                 = $msgId
                    MessageId          = $msgId  # Also add MessageId for compatibility
                    ReceivedDateTime   = $msg.ReceivedDateTime
                    Subject            = $msg.Subject
                    SenderName         = if ($msg.SenderName) { $msg.SenderName } else { "N/A" }
                    SenderEmailAddress = if ($msg.SenderEmailAddress) { $msg.SenderEmailAddress } else { "N/A" }
                    Size               = if ($msg.Size) { $msg.Size } else { 0 }
                    HasAttachments     = if ($msg.HasAttachments) { $msg.HasAttachments } else { $false }
                    BodyPreview        = if ($msg.BodyPreview) { $msg.BodyPreview } else { "" }
                }
            }

            # Display using standardized email list view
            Show-StandardizedEmailListView -UserEmail $UserEmail `
                                           -Messages $messagesForView `
                                           -Title "Emails from: $($selected.Domain) ($($messagesForView.Count) emails)" `
                                           -AllowActions $true `
                                           -ViewName "SenderOverview_$($selected.Domain)"
        } else {
            Write-Host "No emails found for this sender." -ForegroundColor Yellow
            Wait-EnterKey
        }
    }
}

#endregion

#region Main Loop
$mainLoopActive = $true

try {
    while ($mainLoopActive) {
        $selection = Show-MainMenuOptions -UserEmail $MailboxEmail

        switch ($selection.ToUpper()) {
            "1" {
                # Option 1: Sender Overview (from cache)
                Show-SenderOverviewMenu -UserEmail $MailboxEmail
            }
            "2" {
                # Option 2: Manage emails from specific sender
                Show-SenderEmailsMenu -UserEmail $MailboxEmail
            }
            "3" {
                # Option 3: Search for an email (live)
                Invoke-EmailSearch -UserEmail $MailboxEmail
            }
            "4" {
                # Option 4: View last 100 emails (live)
                Show-RecentEmails -UserEmail $MailboxEmail -Count 100
            }
            "5" {
                # Option 5: Empty Deleted Items (live)
                $confirm = Show-Confirmation -Message "Are you sure you want to empty the Deleted Items folder? This cannot be undone."
                if ($confirm) {
                    Write-Host "Emptying Deleted Items folder..." -ForegroundColor Cyan
                    $result = Clear-GraphDeletedItems -UserId $MailboxEmail
                    if ($result) {
                        Write-Host "Deleted Items folder emptied successfully." -ForegroundColor Green
                    } else {
                        Write-Host "Failed to empty Deleted Items folder. See error above." -ForegroundColor Red
                    }
                } else {
                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                }
                Wait-EnterKey
            }
            "A" {
                # Advanced menu: Smart Features
                Show-SmartFeaturesMenu -UserEmail $MailboxEmail
            }
            "B" {
                # Advanced menu: Bulk Operations
                Show-BulkOperationsMenu -UserEmail $MailboxEmail
            }
            "R" {
                Write-Host (Get-LocalizedString "mainMenu_actionStartingFullIndex") -ForegroundColor Cyan
                Clear-SenderCache
                Build-MailboxIndex -UserId $MailboxEmail -MaxEmailsToIndex $MaxEmailsToIndex -TestMode:$TestRun
                Write-Host (Get-LocalizedString "mainMenu_actionIndexingCompleteReload") -ForegroundColor Green
                Wait-EnterKey
            }
            "Q" {
                Write-Host (Get-LocalizedString "mainMenu_actionQuitting") -ForegroundColor Yellow
                $mainLoopActive = $false
            }
            default {
                Write-Host (Get-LocalizedString "mainMenu_actionUnknown" -FormatArgs $selection) -ForegroundColor Red
                Wait-EnterKey
            }
        }
    }
} catch {
    Write-Error (Get-LocalizedString "script_errorOccurred" -FormatArgs $_.Exception.Message)
    Write-Error $_.ScriptStackTrace
} finally {
    # Cleanup
    Write-Host ""
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-GraphService

    # Reset console colors
    Reset-ConsoleColors

    Write-Host ""
    Write-Host "Thank you for using MailCleanBuddy!" -ForegroundColor Green
    Write-Host ""
}
#endregion
