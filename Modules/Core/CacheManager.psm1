<#
.SYNOPSIS
    Cache management module for MailCleanBuddy
.DESCRIPTION
    Handles loading, saving, and updating the local mailbox cache
#>

# Import required modules

$Script:SenderCache = @{}
$Script:CacheFilePath = $null

<#
.SYNOPSIS
    Gets the cache file path for a mailbox
.PARAMETER MailboxEmail
    The mailbox email address
#>
function Get-CacheFilePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxEmail,

        [Parameter(Mandatory = $false)]
        [string]$BasePath
    )

    if ([string]::IsNullOrWhiteSpace($BasePath)) {
        $BasePath = $PSScriptRoot
    }

    $safeEmailForFilename = $MailboxEmail -replace '[^a-zA-Z0-9@._-]', '_'
    $cacheFileName = "mailcleanbuddy_cache_$safeEmailForFilename.json"
    $Script:CacheFilePath = Join-Path -Path $BasePath -ChildPath "..\..\" | Join-Path -ChildPath $cacheFileName

    Write-Verbose "Cache file path: $Script:CacheFilePath"
    return $Script:CacheFilePath
}

<#
.SYNOPSIS
    Loads the local cache from disk
#>
function Import-LocalCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$FilePath
    )

    if ($FilePath) {
        $Script:CacheFilePath = $FilePath
    }

    if ([string]::IsNullOrWhiteSpace($Script:CacheFilePath)) {
        Write-Warning "Cache file path is not set. Cannot load cache."
        return $false
    }

    if (-not (Test-Path $Script:CacheFilePath)) {
        Write-Verbose "No local cache found at: $Script:CacheFilePath"
        return $false
    }

    try {
        Write-Host "Loading local cache from: $Script:CacheFilePath" -ForegroundColor Cyan
        Write-Progress -Activity "Loading Cache" -Status "Reading cache file..." -PercentComplete 0

        $jsonContent = Get-Content -Path $Script:CacheFilePath -Raw -ErrorAction Stop
        Write-Progress -Activity "Loading Cache" -Status "Parsing JSON..." -PercentComplete 30

        $loadedCache = ConvertFrom-Json -InputObject $jsonContent -AsHashtable -ErrorAction Stop
        Write-Progress -Activity "Loading Cache" -Status "Processing messages..." -PercentComplete 60

        # Convert to proper structure
        $Script:SenderCache = @{}
        foreach ($key in $loadedCache.Keys) {
            $cacheEntry = $loadedCache[$key]
            $Script:SenderCache[$key] = @{
                Name = $cacheEntry.Name
                Count = $cacheEntry.Count
                Messages = [System.Collections.Generic.List[PSObject]]::new()
            }
            foreach ($msg in $cacheEntry.Messages) {
                $Script:SenderCache[$key].Messages.Add($msg)
            }
        }

        Write-Progress -Activity "Loading Cache" -Completed
        Write-Host "Local cache loaded successfully. $($Script:SenderCache.Keys.Count) domains found." -ForegroundColor Green
        return $true
    } catch {
        Write-Warning "Error loading cache: $($_.Exception.Message). Cache will be ignored."
        $Script:SenderCache = @{}
        Write-Progress -Activity "Loading Cache" -Completed
        return $false
    }
}

<#
.SYNOPSIS
    Saves the cache to disk
#>
function Export-LocalCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$FilePath
    )

    if ($FilePath) {
        $Script:CacheFilePath = $FilePath
    }

    if ([string]::IsNullOrWhiteSpace($Script:CacheFilePath)) {
        Write-Warning "Cache file path is not set. Cannot save cache."
        return $false
    }

    if ($null -eq $Script:SenderCache -or $Script:SenderCache.Count -eq 0) {
        Write-Warning "SenderCache is empty. Cache will not be saved."
        return $false
    }

    try {
        Write-Host "Saving local cache to: $Script:CacheFilePath" -ForegroundColor Cyan
        $jsonContent = ConvertTo-Json -InputObject $Script:SenderCache -Depth 10 -ErrorAction Stop
        Set-Content -Path $Script:CacheFilePath -Value $jsonContent -ErrorAction Stop
        Write-Host "Local cache saved successfully." -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Error saving cache: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Gets the current sender cache
#>
function Get-SenderCache {
    return $Script:SenderCache
}

<#
.SYNOPSIS
    Sets the sender cache
#>
function Set-SenderCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Cache
    )
    $Script:SenderCache = $Cache
}

<#
.SYNOPSIS
    Updates cache after deleting or moving a message
.PARAMETER Domain
    The sender domain
.PARAMETER MessageId
    The message ID to remove
#>
function Update-CacheAfterAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain,

        [Parameter(Mandatory = $true)]
        [string]$MessageId
    )

    $domainKey = $Domain.ToLowerInvariant()

    if ($Script:SenderCache.ContainsKey($domainKey)) {
        $messageToRemove = $Script:SenderCache[$domainKey].Messages | Where-Object { $_.MessageId -eq $MessageId }
        if ($messageToRemove) {
            $Script:SenderCache[$domainKey].Messages.Remove($messageToRemove)
            $Script:SenderCache[$domainKey].Count = $Script:SenderCache[$domainKey].Messages.Count

            # If no messages left, remove domain from cache
            if ($Script:SenderCache[$domainKey].Count -eq 0) {
                $Script:SenderCache.Remove($domainKey)
                Write-Verbose "Domain '$domainKey' removed from cache (no messages left)"
            }

            Write-Verbose "Cache updated for domain: $domainKey"
            return $true
        }
    }

    Write-Verbose "Message not found in cache: $MessageId"
    return $false
}

<#
.SYNOPSIS
    Clears the entire cache
#>
function Clear-SenderCache {
    $Script:SenderCache = @{}
    Write-Verbose "Cache cleared"
}

<#
.SYNOPSIS
    Indexes the mailbox and populates the cache
#>
function Build-MailboxIndex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $false)]
        [int]$MaxEmailsToIndex = 0,

        [Parameter(Mandatory = $false)]
        [switch]$TestMode
    )

    Write-Host "Starting mailbox indexing for $UserId..." -ForegroundColor Cyan

    if ($MaxEmailsToIndex -gt 0) {
        Write-Warning "MaxEmailsToIndex ACTIVE: Indexing last $MaxEmailsToIndex emails only."
    } elseif ($TestMode) {
        Write-Warning "TEST MODE ACTIVE: Indexing last 100 emails only."
    }

    $Script:SenderCache = @{}

    try {
        # Build parameters for Get-GraphMessages
        $params = @{
            UserId = $UserId
        }

        if ($MaxEmailsToIndex -gt 0) {
            $params.Top = $MaxEmailsToIndex
            $params.OrderBy = "receivedDateTime desc"
            Write-Host "Configuration: Retrieving last $MaxEmailsToIndex messages (incl. MAPI size)."
        } elseif ($TestMode) {
            $params.Top = 100
            $params.OrderBy = "receivedDateTime desc"
            Write-Host "Configuration: Retrieving last 100 messages (TEST MODE)."
        } else {
            $params.All = $true
            Write-Host "Configuration: Retrieving all messages (Full mode). This may take some time."
        }

        Write-Host "Fetching messages..." -ForegroundColor Cyan
        $messages = Get-GraphMessages @params

        if ($null -eq $messages -or $messages.Count -eq 0) {
            Write-Warning "No messages found in mailbox during indexing."
            return $false
        }

        Write-Host "$($messages.Count) messages found. Processing senders..." -ForegroundColor Cyan

        $processedCount = 0
        $totalMessages = $messages.Count
        $updateInterval = [math]::Ceiling($totalMessages / 20)
        if ($updateInterval -eq 0) { $updateInterval = 1 }

        foreach ($message in $messages) {
            $processedCount++
            if ($processedCount % $updateInterval -eq 0 -or $processedCount -eq $totalMessages) {
                Write-Progress -Activity "Indexing Mailbox" -Status "Processing messages..." `
                    -PercentComplete (($processedCount / $totalMessages) * 100) `
                    -CurrentOperation "$processedCount of $totalMessages messages processed"
            }

            $emailSenderAddressInfo = $message.Sender.EmailAddress
            if ($emailSenderAddressInfo -and $emailSenderAddressInfo.Address) {
                $senderFullAddress = $emailSenderAddressInfo.Address
                $domain = ($senderFullAddress -split '@')[1]
                if ([string]::IsNullOrWhiteSpace($domain)) {
                    $domain = "unknown_domain"
                }
                $domainKey = $domain.ToLowerInvariant()

                # Get message size from MAPI properties
                $currentMessageSize = $null
                $messageSizeMapiPropertyId = "Integer 0x0E08"
                $mapiSizeProp = $message.SingleValueExtendedProperties | Where-Object { $_.Id -eq $messageSizeMapiPropertyId } | Select-Object -First 1
                if ($mapiSizeProp -and $mapiSizeProp.Value) {
                    try {
                        $currentMessageSize = [long]$mapiSizeProp.Value
                    } catch {
                        Write-Verbose "Failed to parse MAPI size property for message $($message.Id): $($_.Exception.Message)"
                    }
                }

                # Get attachment flag from MAPI properties
                $currentHasAttachments = $false
                $messageHasAttachMapiPropertyId = "Boolean 0x0E1B"
                $mapiAttachProp = $message.SingleValueExtendedProperties | Where-Object { $_.Id -eq $messageHasAttachMapiPropertyId } | Select-Object -First 1
                if ($mapiAttachProp -and $mapiAttachProp.Value -ne $null) {
                    try {
                        $currentHasAttachments = [System.Convert]::ToBoolean($mapiAttachProp.Value)
                    } catch {
                        Write-Verbose "Failed to parse MAPI attachment flag for message $($message.Id): $($_.Exception.Message)"
                    }
                }

                $messageDetail = @{
                    MessageId = $message.Id
                    Subject = $message.Subject
                    ReceivedDateTime = $message.ReceivedDateTime
                    SenderName = $emailSenderAddressInfo.Name
                    SenderEmailAddress = $senderFullAddress
                    Size = $currentMessageSize
                    HasAttachments = $currentHasAttachments
                    ToRecipients = $message.ToRecipients | ForEach-Object { $_.EmailAddress.Address }
                    Categories = $message.Categories
                }

                if ($Script:SenderCache.ContainsKey($domainKey)) {
                    $Script:SenderCache[$domainKey].Count++
                    $Script:SenderCache[$domainKey].Messages.Add($messageDetail)
                } else {
                    $Script:SenderCache[$domainKey] = @{
                        Name = $domainKey
                        Count = 1
                        Messages = [System.Collections.Generic.List[PSObject]]::new()
                    }
                    $Script:SenderCache[$domainKey].Messages.Add($messageDetail)
                }
            }
        }

        Write-Progress -Activity "Indexing Mailbox" -Completed

        $uniqueSenders = $Script:SenderCache.Keys.Count
        Write-Host "Indexing completed. $uniqueSenders unique sender domain(s) found." -ForegroundColor Green

        # Save cache
        Export-LocalCache

        return $true
    } catch {
        Write-Error "Error during mailbox indexing: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
        }
        return $false
    }
}

Export-ModuleMember -Function Get-CacheFilePath, Import-LocalCache, Export-LocalCache, Get-SenderCache, Set-SenderCache, `
                              Update-CacheAfterAction, Clear-SenderCache, Build-MailboxIndex
