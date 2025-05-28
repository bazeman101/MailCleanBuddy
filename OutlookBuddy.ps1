<#
.SYNOPSIS
    Downloads DMARC reports from a specified mailbox in Microsoft 365.
.DESCRIPTION
    This script connects to Microsoft 365 using interactive login,
    searches for emails containing DMARC reports in the specified mailbox,
    and downloads the report attachments to a local folder.
.PARAMETER MailboxEmail
    The email address of the mailbox to search for DMARC reports.
.PARAMETER ReportsPath
    The local path where the DMARC reports should be saved. Defaults to "_reports" in the script's directory.
.EXAMPLE
    .\Get-DmarcReportsFromM365.ps1 -MailboxEmail "dmarc-reports@example.com"
    This command will connect to the "dmarc-reports@example.com" mailbox and save reports to ".\_reports".
.EXAMPLE
    .\Get-DmarcReportsFromM365.ps1 -MailboxEmail "dmarc-reports@example.com" -ReportsPath "C:\DMARC_Reports"
    This command will connect to the "dmarc-reports@example.com" mailbox and save reports to "C:\DMARC_Reports".
.NOTES
    Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Mail modules.
    The script will attempt to install them if not found.
    Ensure you have the necessary permissions (Microsoft Graph: Mail.Read) to access the specified mailbox.
    You will be prompted to consent to these permissions on first run.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$MailboxEmail,

    [Parameter(Mandatory = $false)]
    [string]$ReportsPath = "_reports"
)

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

    # Resolve and create the reports path if it doesn't exist
    $ResolvedReportsPath = Join-Path -Path $PSScriptRoot -ChildPath $ReportsPath
    if (-not (Test-Path -Path $ResolvedReportsPath)) {
        Write-Host "Creating report directory: $ResolvedReportsPath"
        New-Item -ItemType Directory -Path $ResolvedReportsPath -Force | Out-Null
    } else {
        Write-Host "Report directory already exists: $ResolvedReportsPath"
    }

    # Search for emails with DMARC reports.
    # DMARC reports often have subjects like "Report Domain: <domain> Submitter: <submitter>"
    # and attachments are typically .xml, .gz, or .zip files.
    # We'll search for subjects containing "Report Domain:" as a common indicator.
    # Adjust the search query as needed for your specific report format.
    # Using Microsoft Graph $filter syntax.
    $filterQuery = "contains(subject, 'Report Domain:')"
    Write-Host "Attempting to access Inbox for mailbox '$MailboxEmail'..."
    try {
        # Get the Inbox folder specifically. "inbox" is a well-known folder name.
        $inboxFolder = Get-MgUserMailFolder -UserId $MailboxEmail -MailFolderId "inbox" -ErrorAction Stop
        Write-Host "Successfully accessed Inbox for mailbox '$MailboxEmail' (Folder ID: $($inboxFolder.Id))."
    }
    catch {
        throw "Kritiek: Kon de Inbox folder niet benaderen voor mailbox '$MailboxEmail'. Controleer de mailboxnaam en rechten. Foutdetails: $($_.Exception.Message)"
    }
    
    Write-Host "Searching for DMARC report emails in Inbox of '$MailboxEmail' with filter: $filterQuery"
    # Using Get-MgUserMailFolderMessage to find messages in the specific folder (Inbox).
    # The -All parameter handles pagination to retrieve all matching messages.
    $messages = Get-MgUserMailFolderMessage -UserId $MailboxEmail -MailFolderId $inboxFolder.Id -Filter $filterQuery -All -ErrorAction Stop

    if ($messages.Count -eq 0) {
        Write-Host "No DMARC report emails found matching the criteria."
    } else {
        Write-Host "$($messages.Count) DMARC report email(s) found."

        foreach ($message in $messages) {
            Write-Host "Processing email: $($message.Subject) (Received: $($message.ReceivedDateTime))"
            
            # Get attachments for the current message using Microsoft Graph
            # We need to ensure attachments are actually present before trying to get them.
            if ($message.HasAttachments) {
                # Remove -ErrorAction SilentlyContinue to see if there are errors fetching attachments
                $attachments = Get-MgUserMessageAttachment -UserId $MailboxEmail -MessageId $message.Id -ErrorAction Stop 
                
                if ($attachments.Count -gt 0) {
                    foreach ($attachment in $attachments) {
                        Write-Host "DEBUG: Found attachment - Name: $($attachment.Name), ODataType: $($attachment.OdataType), Size: $($attachment.Size), ContentId: $($attachment.ContentId)"
                        # Filter for common DMARC report file types.
                        # We rely on the filename and the subsequent check for ContentBytes.
                        if ($attachment.Name -like "*.xml" -or $attachment.Name -like "*.xml.gz" -or $attachment.Name -like "*.zip" -or $attachment.Name -like "*.tar" -or $attachment.Name -like "*.tar.gz") {
                            
                            $originalFilePath = Join-Path -Path $ResolvedReportsPath -ChildPath $attachment.Name
                            
                            # Check if the file has already been downloaded
                            if (Test-Path $originalFilePath) {
                                Write-Host "Skipping already downloaded attachment: $($attachment.Name)"
                                continue # Skip to the next attachment
                            }

                            # File does not exist, proceed to save.
                            $filePath = $originalFilePath
                            $counter = 1
                            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($attachment.Name)
                            $extension = [System.IO.Path]::GetExtension($attachment.Name)
                            # This loop handles rare cases where a new file might conflict if downloaded in the same run with same name
                            while (Test-Path $filePath) { 
                                $newFileName = "{0}_{1}{2}" -f $baseName, $counter, $extension
                                $filePath = Join-Path -Path $ResolvedReportsPath -ChildPath $newFileName
                                $counter++
                            }

                            Write-Host "Attempting to save attachment: $($attachment.Name) to $filePath"
                            
                            # Use Invoke-MgGraphRequest to get the raw content of the attachment ($value endpoint) and save to file
                            $attachmentValueUri = "/users/$MailboxEmail/messages/$($message.Id)/attachments/$($attachment.Id)/`$value" # Note: $value needs escaping with backtick for PowerShell
                            try {
                                # Get the raw content of the attachment
                                $attachmentContent = Invoke-MgGraphRequest -Method GET -Uri $attachmentValueUri -ErrorAction Stop
                                
                                if ($attachmentContent) {
                                    [System.IO.File]::WriteAllBytes($filePath, $attachmentContent)
                                    Write-Host "Successfully saved attachment: $($attachment.Name) to $filePath"
                                } else {
                                    Write-Warning "Invoke-MgGraphRequest returned no content for attachment '$($attachment.Name)'. Skipping."
                                }
                            }
                            catch {
                                Write-Warning "Failed to retrieve or save attachment '$($attachment.Name)' to '$filePath' using Invoke-MgGraphRequest. Error: $($_.Exception.Message). Skipping."
                                # If the file was partially created before an error, attempt to remove it.
                                if (Test-Path $filePath) {
                                    Remove-Item $filePath -ErrorAction SilentlyContinue
                                }
                                continue # Skip to the next attachment
                            }
                        } else {
                            # This message now means the filename pattern did not match.
                            Write-Host "Skipping attachment '$($attachment.Name)' as its name does not match DMARC report patterns."
                        }
                    }
                } else {
                    Write-Host "No attachments found for email: $($message.Subject) despite HasAttachments being true, or error fetching attachments."
                }
            } else {
                Write-Host "Email '$($message.Subject)' has no attachments indicated."
            }
        }
    }

}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
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
