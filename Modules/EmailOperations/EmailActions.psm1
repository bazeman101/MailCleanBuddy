<#
.SYNOPSIS
    Email actions module for MailCleanBuddy
.DESCRIPTION
    Handles email operations: delete, move, bulk operations, and sender management
#>

# Import required modules

<#
.SYNOPSIS
    Deletes multiple emails
.PARAMETER UserId
    User email address
.PARAMETER MessageIds
    Array of message IDs to delete
#>
function Remove-BulkEmails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string[]]$MessageIds
    )

    $deletedCount = 0
    $errorCount = 0
    $total = $MessageIds.Count

    Write-Host "`nDeleting $total emails..." -ForegroundColor Cyan

    $index = 0
    foreach ($msgId in $MessageIds) {
        $index++
        Write-Progress -Activity "Deleting Emails" -Status "Deleting email $index of $total" -PercentComplete (($index / $total) * 100)

        try {
            $result = Remove-GraphMessage -UserId $UserId -MessageId $msgId
            if ($result) {
                $deletedCount++
            } else {
                $errorCount++
            }
        } catch {
            Write-Warning "Error deleting message $msgId - $($_.Exception.Message)"
            $errorCount++
        }
    }

    Write-Progress -Activity "Deleting Emails" -Completed

    Write-Host "`nDeleted: $deletedCount of $total emails" -ForegroundColor Green
    if ($errorCount -gt 0) {
        Write-Host "Errors: $errorCount" -ForegroundColor Red
    }

    return @{
        DeletedCount = $deletedCount
        ErrorCount = $errorCount
        Total = $total
    }
}

<#
.SYNOPSIS
    Moves multiple emails to a folder
.PARAMETER UserId
    User email address
.PARAMETER MessageIds
    Array of message IDs to move
.PARAMETER DestinationFolderId
    Destination folder ID
#>
function Move-BulkEmails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string[]]$MessageIds,

        [Parameter(Mandatory = $true)]
        [string]$DestinationFolderId
    )

    $movedCount = 0
    $errorCount = 0
    $total = $MessageIds.Count

    Write-Host "`nMoving $total emails..." -ForegroundColor Cyan

    $index = 0
    foreach ($msgId in $MessageIds) {
        $index++
        Write-Progress -Activity "Moving Emails" -Status "Moving email $index of $total" -PercentComplete (($index / $total) * 100)

        try {
            $result = Move-GraphMessage -UserId $UserId -MessageId $msgId -DestinationFolderId $DestinationFolderId
            if ($result) {
                $movedCount++
            } else {
                $errorCount++
            }
        } catch {
            Write-Warning "Error moving message $msgId - $($_.Exception.Message)"
            $errorCount++
        }
    }

    Write-Progress -Activity "Moving Emails" -Completed

    Write-Host "`nMoved: $movedCount of $total emails" -ForegroundColor Green
    if ($errorCount -gt 0) {
        Write-Host "Errors: $errorCount" -ForegroundColor Red
    }

    return @{
        MovedCount = $movedCount
        ErrorCount = $errorCount
        Total = $total
    }
}

<#
.SYNOPSIS
    Moves all emails from a sender domain to a subfolder (creates if needed)
.PARAMETER UserId
    User email address
.PARAMETER SenderDomain
    Sender domain to filter
.PARAMETER SubfolderName
    Name of subfolder to create/use
.PARAMETER ParentFolderId
    Parent folder ID (default: inbox)
.PARAMETER PreviewOnly
    Only show what would be moved without actually moving
#>
function Move-SenderToSubfolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$SenderDomain,

        [Parameter(Mandatory = $true)]
        [string]$SubfolderName,

        [Parameter(Mandatory = $false)]
        [string]$ParentFolderId = "inbox",

        [Parameter(Mandatory = $false)]
        [switch]$PreviewOnly
    )

    Write-Host "`n=== Move Sender to Subfolder ===" -ForegroundColor Cyan
    Write-Host "Sender Domain: $SenderDomain" -ForegroundColor Yellow
    Write-Host "Subfolder: $SubfolderName" -ForegroundColor Yellow
    if ($PreviewOnly) {
        Write-Host "MODE: PREVIEW ONLY (no changes will be made)" -ForegroundColor Yellow
    }

    try {
        # Get or create subfolder
        $targetFolder = $null
        Write-Host "`nChecking for existing subfolder..." -ForegroundColor Cyan

        # Find parent folder - handle well-known folder names
        $parentFolder = $null

        # List of well-known folder names that can be used directly as IDs
        $wellKnownFolders = @("inbox", "sentitems", "deleteditems", "drafts", "junkemail", "archive")

        if ($wellKnownFolders -contains $ParentFolderId.ToLower()) {
            # For well-known folders, fetch directly using the well-known name
            try {
                $parentFolder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $ParentFolderId -ErrorAction Stop
                Write-Host "Using well-known folder: $($parentFolder.DisplayName)" -ForegroundColor Cyan
            } catch {
                Write-Verbose "Could not fetch well-known folder '$ParentFolderId', searching by display name..."
            }
        }

        # If not found yet, search all folders
        if (-not $parentFolder) {
            $allFolders = Get-GraphMailFolders -UserId $UserId
            # Case-insensitive comparison for DisplayName
            $parentFolder = $allFolders | Where-Object {
                $_.Id -eq $ParentFolderId -or
                $_.DisplayName -eq $ParentFolderId -or
                $_.DisplayName.ToLower() -eq $ParentFolderId.ToLower()
            } | Select-Object -First 1
        }

        if (-not $parentFolder) {
            Write-Error "Parent folder not found: $ParentFolderId"
            return
        }

        Write-Host "Parent folder: $($parentFolder.DisplayName)" -ForegroundColor Green

        # Get child folders of parent
        $childFolders = Get-MgUserMailFolderChildFolder -UserId $UserId -MailFolderId $parentFolder.Id -All -ErrorAction Stop

        # Check if subfolder exists
        $targetFolder = $childFolders | Where-Object { $_.DisplayName -eq $SubfolderName } | Select-Object -First 1

        if ($null -eq $targetFolder) {
            if ($PreviewOnly) {
                Write-Host "Subfolder '$SubfolderName' does not exist (would be created)" -ForegroundColor Yellow
                $targetFolderId = "PREVIEW_MODE_FOLDER_ID"
            } else {
                Write-Host "Creating subfolder '$SubfolderName'..." -ForegroundColor Cyan
                $targetFolder = New-GraphMailFolder -UserId $UserId -DisplayName $SubfolderName -ParentFolderId $parentFolder.Id
                Write-Host "Created subfolder: $($targetFolder.DisplayName)" -ForegroundColor Green
                $targetFolderId = $targetFolder.Id
            }
        } else {
            Write-Host "Found existing subfolder: $($targetFolder.DisplayName)" -ForegroundColor Green
            $targetFolderId = $targetFolder.Id
        }

        # Find all emails from sender domain
        Write-Host "`nSearching for emails from domain: $SenderDomain..." -ForegroundColor Cyan

        $filter = "contains(from/emailAddress/address, '@" + $SenderDomain + "')"
        $messages = Get-GraphMessages -UserId $UserId -Filter $filter -All

        if ($null -eq $messages -or $messages.Count -eq 0) {
            Write-Warning "No emails found from domain: $SenderDomain"
            return
        }

        Write-Host "Found $($messages.Count) emails from $SenderDomain" -ForegroundColor Green

        if ($PreviewOnly) {
            # Preview mode - show what would be moved
            Write-Host "`n=== PREVIEW: Emails that would be moved ===" -ForegroundColor Yellow
            $previewCount = [Math]::Min(10, $messages.Count)
            for ($i = 0; $i -lt $previewCount; $i++) {
                $msg = $messages[$i]
                Write-Host "  [$($i+1)] $($msg.ReceivedDateTime.ToString('yyyy-MM-dd')) - $($msg.Subject)" -ForegroundColor White
            }
            if ($messages.Count -gt 10) {
                Write-Host "  ... and $($messages.Count - 10) more emails" -ForegroundColor White
            }

            Write-Host "`nTotal emails that would be moved: $($messages.Count)" -ForegroundColor Yellow
            Write-Host "Run without -PreviewOnly to actually move these emails" -ForegroundColor Yellow
            return
        }

        # Confirm action
        Write-Host "`nThis will move $($messages.Count) emails to subfolder '$SubfolderName'" -ForegroundColor Yellow
        $confirmation = Read-Host "Continue? (yes/no)"
        if ($confirmation -notmatch '^(y|yes)$') {
            Write-Host "Operation cancelled" -ForegroundColor Yellow
            return
        }

        # Move emails
        $messageIds = $messages | ForEach-Object { $_.Id }
        $result = Move-BulkEmails -UserId $UserId -MessageIds $messageIds -DestinationFolderId $targetFolderId

        # Summary
        Write-Host "`n=== Summary ===" -ForegroundColor Cyan
        Write-Host "Sender Domain: $SenderDomain" -ForegroundColor White
        Write-Host "Destination Folder: $SubfolderName" -ForegroundColor White
        Write-Host "Total Moved: $($result.MovedCount) of $($result.Total)" -ForegroundColor Green
        if ($result.ErrorCount -gt 0) {
            Write-Host "Errors: $($result.ErrorCount)" -ForegroundColor Red
        }

    } catch {
        Write-Error "Error moving sender to subfolder: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Gets emails from a specific sender domain
.PARAMETER UserId
    User email address
.PARAMETER SenderDomain
    Sender domain to filter
.PARAMETER IncludeCount
    Only return count, not full messages
#>
function Get-EmailsBySenderDomain {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$SenderDomain,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeCount
    )

    try {
        $filter = "contains(from/emailAddress/address, '@" + $SenderDomain + "')"
        $messages = Get-GraphMessages -UserId $UserId -Filter $filter -All

        if ($IncludeCount) {
            return @{
                Count = if ($messages) { $messages.Count } else { 0 }
                Domain = $SenderDomain
            }
        }

        return $messages

    } catch {
        Write-Error "Error retrieving emails by sender domain: $($_.Exception.Message)"
        return $null
    }
}

<#
.SYNOPSIS
    Interactive folder selection menu
.PARAMETER UserId
    User email address
#>
function Select-MailFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        $folders = Get-GraphMailFolders -UserId $UserId

        if ($null -eq $folders -or $folders.Count -eq 0) {
            Write-Warning "No folders found"
            return $null
        }

        # Sort folders by display name
        $sortedFolders = $folders | Sort-Object DisplayName

        Write-Host "`n=== Select Mail Folder ===" -ForegroundColor Cyan
        for ($i = 0; $i -lt $sortedFolders.Count; $i++) {
            Write-Host "  [$($i+1)] $($sortedFolders[$i].DisplayName)" -ForegroundColor Green
        }
        Write-Host "  [0] Cancel" -ForegroundColor Yellow

        do {
            $selection = Read-Host "`nEnter folder number"
            $selectionInt = 0
            $validSelection = [int]::TryParse($selection, [ref]$selectionInt)
        } while (-not $validSelection -or $selectionInt -lt 0 -or $selectionInt -gt $sortedFolders.Count)

        if ($selectionInt -eq 0) {
            return $null
        }

        return $sortedFolders[$selectionInt - 1]

    } catch {
        Write-Error "Error selecting folder: $($_.Exception.Message)"
        return $null
    }
}

Export-ModuleMember -Function Remove-BulkEmails, Move-BulkEmails, Move-SenderToSubfolder, Get-EmailsBySenderDomain, Select-MailFolder
