<#
.SYNOPSIS
    Attachment management module for MailCleanBuddy
.DESCRIPTION
    Handles downloading attachments from emails using the CORRECT Microsoft Graph API method
    Uses AdditionalProperties.contentBytes with base64 decoding
#>

# Import required modules

<#
.SYNOPSIS
    Downloads a single attachment using CORRECT base64 method
.PARAMETER UserId
    User email address
.PARAMETER MessageId
    Message ID
.PARAMETER AttachmentId
    Attachment ID
.PARAMETER SavePath
    Path to save the attachment
.PARAMETER SenderEmail
    Sender email for filename generation
.PARAMETER ReceivedDate
    Date received for filename generation
.PARAMETER OriginalFilename
    Original attachment filename
#>
function Save-EmailAttachment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$MessageId,

        [Parameter(Mandatory = $true)]
        [string]$AttachmentId,

        [Parameter(Mandatory = $true)]
        [string]$SavePath,

        [Parameter(Mandatory = $false)]
        [string]$SenderEmail,

        [Parameter(Mandatory = $false)]
        [DateTime]$ReceivedDate,

        [Parameter(Mandatory = $false)]
        [string]$OriginalFilename
    )

    try {
        # Get attachment using Microsoft Graph
        Write-Verbose "Retrieving attachment $AttachmentId from message $MessageId"
        $attachment = Get-MgUserMessageAttachment -UserId $UserId -MessageId $MessageId -AttachmentId $AttachmentId -ErrorAction Stop

        # Check attachment type
        $attachmentType = $attachment.AdditionalProperties.'@odata.type'
        Write-Verbose "Attachment type: $attachmentType"

        if ($attachmentType -ne '#microsoft.graph.fileAttachment') {
            Write-Warning "Attachment is not a file attachment (type: $attachmentType). Skipping."
            return $null
        }

        # CORRECT METHOD: Get contentBytes from AdditionalProperties
        $base64Content = $attachment.AdditionalProperties.contentBytes

        if ([string]::IsNullOrWhiteSpace($base64Content)) {
            Write-Warning "Attachment has no content. Skipping."
            return $null
        }

        # Decode base64 to bytes
        Write-Verbose "Decoding base64 content..."
        $bytes = [System.Convert]::FromBase64String($base64Content)

        if ($bytes.Length -eq 0) {
            Write-Warning "Decoded attachment is empty. Skipping."
            return $null
        }

        # Generate filename using new convention: yy-MM-dd verzender - originele_bestandsnaam.ext
        $filename = $OriginalFilename
        if ($ReceivedDate -and $SenderEmail) {
            $dateStr = $ReceivedDate.ToString("yy-MM-dd")
            $senderPart = ($SenderEmail -split '@')[0]
            $senderPart = Get-SafeFilename -Text $senderPart -MaxLength 30
            $filename = "$dateStr $senderPart - $OriginalFilename"
        }

        # Ensure unique filename
        $fullPath = Join-Path -Path $SavePath -ChildPath $filename
        $fullPath = Get-UniqueFilePath -FilePath $fullPath

        # Write bytes to file
        Write-Verbose "Writing $($bytes.Length) bytes to $fullPath"
        [System.IO.File]::WriteAllBytes($fullPath, $bytes)

        Write-Host "Downloaded: $filename ($(Format-FileSize -Bytes $bytes.Length))" -ForegroundColor Green

        return @{
            Success = $true
            FilePath = $fullPath
            FileName = $filename
            Size = $bytes.Length
        }

    } catch {
        Write-Error "Error downloading attachment: $($_.Exception.Message)"
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

<#
.SYNOPSIS
    Downloads all attachments from a single email
.PARAMETER UserId
    User email address
.PARAMETER MessageId
    Message ID
.PARAMETER SavePath
    Path to save attachments
#>
function Get-MessageAttachments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$MessageId,

        [Parameter(Mandatory = $true)]
        [string]$SavePath
    )

    try {
        # Ensure directory exists
        if (-not (Ensure-DirectoryExists -Path $SavePath)) {
            return $null
        }

        # Get message details
        $message = Get-MgUserMessage -UserId $UserId -MessageId $MessageId -ErrorAction Stop

        # Get all attachments
        $attachments = Get-MgUserMessageAttachment -UserId $UserId -MessageId $MessageId -ErrorAction Stop

        if ($null -eq $attachments -or $attachments.Count -eq 0) {
            Write-Warning "No attachments found in message"
            return $null
        }

        $results = @()
        $downloadCount = 0

        foreach ($attachment in $attachments) {
            Write-Host "Processing: $($attachment.Name)..." -ForegroundColor Cyan

            $result = Save-EmailAttachment `
                -UserId $UserId `
                -MessageId $MessageId `
                -AttachmentId $attachment.Id `
                -SavePath $SavePath `
                -SenderEmail $message.From.EmailAddress.Address `
                -ReceivedDate $message.ReceivedDateTime `
                -OriginalFilename $attachment.Name

            if ($result -and $result.Success) {
                $downloadCount++
            }

            $results += $result
        }

        Write-Host "`nDownloaded $downloadCount of $($attachments.Count) attachments" -ForegroundColor Green
        return $results

    } catch {
        Write-Error "Error retrieving attachments: $($_.Exception.Message)"
        return $null
    }
}

<#
.SYNOPSIS
    Bulk downloads attachments from multiple emails with time filters
.PARAMETER UserId
    User email address
.PARAMETER TimeFilter
    Time filter: 'LastDay', 'Last7Days', 'Last30Days', 'Last90Days', 'LastWeek', 'LastMonth', 'All'
.PARAMETER SavePath
    Path to save attachments
.PARAMETER SenderDomain
    Optional: Filter by sender domain
.PARAMETER FileTypes
    Optional: Array of file extensions to filter (e.g., @('pdf', 'xlsx'))
.PARAMETER SkipDuplicates
    Skip files that already exist
#>
function Get-BulkAttachments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [ValidateSet('LastDay', 'Last7Days', 'Last30Days', 'Last90Days', 'LastWeek', 'LastMonth', 'All')]
        [string]$TimeFilter,

        [Parameter(Mandatory = $true)]
        [string]$SavePath,

        [Parameter(Mandatory = $false)]
        [string]$SenderDomain,

        [Parameter(Mandatory = $false)]
        [string[]]$FileTypes,

        [Parameter(Mandatory = $false)]
        [switch]$SkipDuplicates
    )

    Write-Host "`n=== Bulk Attachment Download ===" -ForegroundColor Cyan
    Write-Host "Time Filter: $TimeFilter" -ForegroundColor Yellow
    Write-Host "Save Path: $SavePath" -ForegroundColor Yellow

    # Ensure directory exists
    if (-not (Ensure-DirectoryExists -Path $SavePath)) {
        return
    }

    # Calculate date filter
    $filterDate = $null
    switch ($TimeFilter) {
        'LastDay' { $filterDate = (Get-Date).AddDays(-1) }
        'Last7Days' { $filterDate = (Get-Date).AddDays(-7) }
        'Last30Days' { $filterDate = (Get-Date).AddDays(-30) }
        'Last90Days' { $filterDate = (Get-Date).AddDays(-90) }
        'LastWeek' { $filterDate = (Get-Date).AddDays(-7) }
        'LastMonth' { $filterDate = (Get-Date).AddMonths(-1) }
        'All' { $filterDate = $null }
    }

    # Build OData filter
    $filters = @()
    if ($filterDate) {
        $filterDateStr = $filterDate.ToString('yyyy-MM-ddTHH:mm:ssZ')
        $filters += "receivedDateTime ge $filterDateStr"
    }
    if ($SenderDomain) {
        $filters += "contains(from/emailAddress/address, '$SenderDomain')"
    }
    # Only get emails with attachments
    $filters += "hasAttachments eq true"

    $odataFilter = $filters -join ' and '

    try {
        Write-Host "`nFetching messages with attachments..." -ForegroundColor Cyan

        $messages = Get-GraphMessages -UserId $UserId -Filter $odataFilter -All

        if ($null -eq $messages -or $messages.Count -eq 0) {
            Write-Warning "No messages found matching criteria"
            return
        }

        Write-Host "Found $($messages.Count) messages with attachments" -ForegroundColor Green

        $totalDownloaded = 0
        $totalSkipped = 0
        $totalErrors = 0
        $processedMessages = 0

        foreach ($message in $messages) {
            $processedMessages++
            Write-Progress -Activity "Downloading Attachments" `
                -Status "Processing message $processedMessages of $($messages.Count)" `
                -PercentComplete (($processedMessages / $messages.Count) * 100)

            Write-Host "`n[$processedMessages/$($messages.Count)] Processing: $($message.Subject)" -ForegroundColor Cyan

            try {
                $attachments = Get-MgUserMessageAttachment -UserId $UserId -MessageId $message.Id -ErrorAction Stop

                foreach ($attachment in $attachments) {
                    # Check file type filter
                    if ($FileTypes -and $FileTypes.Count -gt 0) {
                        $extension = [System.IO.Path]::GetExtension($attachment.Name).TrimStart('.')
                        if ($extension -notin $FileTypes) {
                            Write-Verbose "Skipping $($attachment.Name) - file type not in filter"
                            $totalSkipped++
                            continue
                        }
                    }

                    # Generate filename
                    $dateStr = $message.ReceivedDateTime.ToString("yy-MM-dd")
                    $senderPart = ($message.From.EmailAddress.Address -split '@')[0]
                    $senderPart = Get-SafeFilename -Text $senderPart -MaxLength 30
                    $safeOriginalName = Get-SafeFilename -Text $attachment.Name -MaxLength 100
                    $filename = "$dateStr $senderPart - $safeOriginalName"
                    $fullPath = Join-Path -Path $SavePath -ChildPath $filename

                    # Check if file exists and skip if requested
                    if ($SkipDuplicates -and (Test-Path $fullPath)) {
                        Write-Host "  Skipped (exists): $filename" -ForegroundColor Yellow
                        $totalSkipped++
                        continue
                    }

                    # Download attachment
                    $result = Save-EmailAttachment `
                        -UserId $UserId `
                        -MessageId $message.Id `
                        -AttachmentId $attachment.Id `
                        -SavePath $SavePath `
                        -SenderEmail $message.From.EmailAddress.Address `
                        -ReceivedDate $message.ReceivedDateTime `
                        -OriginalFilename $attachment.Name

                    if ($result -and $result.Success) {
                        $totalDownloaded++
                    } else {
                        $totalErrors++
                    }
                }

            } catch {
                Write-Warning "Error processing message: $($_.Exception.Message)"
                $totalErrors++
            }
        }

        Write-Progress -Activity "Downloading Attachments" -Completed

        # Summary
        Write-Host "`n=== Download Summary ===" -ForegroundColor Cyan
        Write-Host "Total Downloaded: $totalDownloaded" -ForegroundColor Green
        Write-Host "Total Skipped: $totalSkipped" -ForegroundColor Yellow
        Write-Host "Total Errors: $totalErrors" -ForegroundColor Red
        Write-Host "Save Location: $SavePath" -ForegroundColor Cyan

    } catch {
        Write-Error "Error during bulk download: $($_.Exception.Message)"
        Write-Progress -Activity "Downloading Attachments" -Completed
    }
}

Export-ModuleMember -Function Save-EmailAttachment, Get-MessageAttachments, Get-BulkAttachments
