<#
.SYNOPSIS
    Microsoft Graph API service module for MailCleanBuddy
.DESCRIPTION
    Handles all Microsoft Graph API interactions including authentication and email operations
#>

$Script:GraphConnected = $false

<#
.SYNOPSIS
    Checks if required Microsoft Graph modules are installed
#>
function Test-GraphModules {
    [CmdletBinding()]
    param()

    $requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Mail")
    $missingModules = @()

    foreach ($moduleName in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            $missingModules += $moduleName
        }
    }

    return @{
        AllInstalled = ($missingModules.Count -eq 0)
        MissingModules = $missingModules
    }
}

<#
.SYNOPSIS
    Installs required Microsoft Graph modules
#>
function Install-GraphModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Modules
    )

    foreach ($module in $Modules) {
        Write-Host "Installing module: $module..." -ForegroundColor Yellow
        try {
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "Successfully installed: $module" -ForegroundColor Green
        } catch {
            Write-Error "Failed to install module '$module': $($_.Exception.Message)"
            throw
        }
    }
}

<#
.SYNOPSIS
    Connects to Microsoft Graph with required scopes
.PARAMETER Scopes
    Array of required permission scopes
#>
function Connect-GraphService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$Scopes = @("Mail.Read", "Mail.ReadWrite", "User.Read")
    )

    try {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $Scopes -ErrorAction Stop -NoWelcome
        $Script:GraphConnected = $true
        Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        $Script:GraphConnected = $false
        return $false
    }
}

<#
.SYNOPSIS
    Disconnects from Microsoft Graph
#>
function Disconnect-GraphService {
    [CmdletBinding()]
    param()

    if ($Script:GraphConnected) {
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            $Script:GraphConnected = $false
            Write-Verbose "Disconnected from Microsoft Graph"
        } catch {
            Write-Warning "Error disconnecting from Microsoft Graph: $($_.Exception.Message)"
        }
    }
}

<#
.SYNOPSIS
    Gets messages from user mailbox
.PARAMETER UserId
    User email address
.PARAMETER Top
    Number of messages to retrieve
.PARAMETER OrderBy
    Order by field
.PARAMETER Filter
    OData filter string
.PARAMETER All
    Retrieve all messages
#>
function Get-GraphMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $false)]
        [int]$Top,

        [Parameter(Mandatory = $false)]
        [string]$OrderBy,

        [Parameter(Mandatory = $false)]
        [string]$Filter,

        [Parameter(Mandatory = $false)]
        [switch]$All
    )

    $params = @{
        UserId = $UserId
        Property = "id,subject,sender,receivedDateTime,toRecipients,categories,from"
        ErrorAction = "Stop"
    }

    # Add extended properties for size and attachments
    $messageSizeMapiPropertyId = "Integer 0x0E08"
    $messageHasAttachMapiPropertyId = "Boolean 0x0E1B"
    $expandExtendedProperties = "singleValueExtendedProperties(`$filter=id eq '$messageSizeMapiPropertyId' or id eq '$messageHasAttachMapiPropertyId')"
    $params.Expand = $expandExtendedProperties

    if ($Top) { $params.Top = $Top }
    if ($OrderBy) { $params.OrderBy = $OrderBy }
    if ($Filter) { $params.Filter = $Filter }
    if ($All) { $params.All = $true }

    try {
        return Get-MgUserMessage @params
    } catch {
        Write-Error "Error retrieving messages: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Gets a specific message by ID
#>
function Get-GraphMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$MessageId
    )

    try {
        return Get-MgUserMessage -UserId $UserId -MessageId $MessageId -ErrorAction Stop
    } catch {
        Write-Error "Error retrieving message: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Deletes a message
#>
function Remove-GraphMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$MessageId
    )

    try {
        Remove-MgUserMessage -UserId $UserId -MessageId $MessageId -ErrorAction Stop
        return $true
    } catch {
        Write-Error "Error deleting message: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Moves a message to a folder
#>
function Move-GraphMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$MessageId,

        [Parameter(Mandatory = $true)]
        [string]$DestinationFolderId
    )

    try {
        $body = @{
            DestinationId = $DestinationFolderId
        }
        Move-MgUserMessage -UserId $UserId -MessageId $MessageId -DestinationId $DestinationFolderId -ErrorAction Stop
        return $true
    } catch {
        Write-Error "Error moving message: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Gets mail folders
#>
function Get-GraphMailFolders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeChildFolders
    )

    try {
        $folders = Get-MgUserMailFolder -UserId $UserId -All -ErrorAction Stop
        return $folders
    } catch {
        Write-Error "Error retrieving mail folders: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Creates a new mail folder
#>
function New-GraphMailFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $false)]
        [string]$ParentFolderId
    )

    try {
        $params = @{
            DisplayName = $DisplayName
        }

        if ($ParentFolderId) {
            $folder = New-MgUserMailFolderChildFolder -UserId $UserId -MailFolderId $ParentFolderId -BodyParameter $params -ErrorAction Stop
        } else {
            $folder = New-MgUserMailFolder -UserId $UserId -BodyParameter $params -ErrorAction Stop
        }

        return $folder
    } catch {
        Write-Error "Error creating mail folder: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Searches messages
#>
function Search-GraphMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$SearchTerm,

        [Parameter(Mandatory = $false)]
        [int]$Top = 100
    )

    try {
        $filter = "contains(subject,'$SearchTerm') or contains(from/emailAddress/address,'$SearchTerm')"
        return Get-GraphMessages -UserId $UserId -Filter $filter -Top $Top
    } catch {
        Write-Error "Error searching messages: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Empties deleted items folder
#>
function Clear-GraphDeletedItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        Write-Host "Getting Deleted Items folder..." -ForegroundColor Cyan
        $deletedItemsFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId "deleteditems" -ErrorAction Stop

        if (-not $deletedItemsFolder) {
            Write-Warning "Could not find Deleted Items folder."
            return $false
        }

        Write-Host "Fetching messages from Deleted Items..." -ForegroundColor Cyan
        $deletedMessages = Get-MgUserMailFolderMessage -UserId $UserId -MailFolderId $deletedItemsFolder.Id -Top 999 -ErrorAction Stop

        if (-not $deletedMessages -or $deletedMessages.Count -eq 0) {
            Write-Host "Deleted Items folder is already empty." -ForegroundColor Green
            return $true
        }

        Write-Host "Found $($deletedMessages.Count) message(s) to delete..." -ForegroundColor Cyan

        $deletedCount = 0
        $errorCount = 0
        $totalCount = $deletedMessages.Count

        foreach ($msg in $deletedMessages) {
            $deletedCount++
            Write-Progress -Activity "Emptying Deleted Items" `
                          -Status "Deleting message $deletedCount of $totalCount" `
                          -PercentComplete (($deletedCount / $totalCount) * 100)

            try {
                Remove-MgUserMessage -UserId $UserId -MessageId $msg.Id -ErrorAction Stop
            } catch {
                $errorCount++
                Write-Verbose "Failed to delete message: $($_.Exception.Message)"
            }
        }

        Write-Progress -Activity "Emptying Deleted Items" -Completed

        $successCount = $deletedCount - $errorCount
        Write-Host "Successfully deleted $successCount of $totalCount message(s)." -ForegroundColor Green

        if ($errorCount -gt 0) {
            Write-Warning "$errorCount message(s) could not be deleted."
        }

        return ($errorCount -eq 0)
    } catch {
        Write-Error "Error emptying deleted items: $($_.Exception.Message)"
        return $false
    }
}

Export-ModuleMember -Function Test-GraphModules, Install-GraphModules, Connect-GraphService, Disconnect-GraphService, `
                              Get-GraphMessages, Get-GraphMessage, Remove-GraphMessage, Move-GraphMessage, `
                              Get-GraphMailFolders, New-GraphMailFolder, Search-GraphMessages, Clear-GraphDeletedItems
