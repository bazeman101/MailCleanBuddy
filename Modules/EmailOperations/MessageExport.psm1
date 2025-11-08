<#
.SYNOPSIS
    Message export module for MailCleanBuddy
.DESCRIPTION
    Exports emails to EML format (MIME) and MSG format (Outlook compatible)
    Note: True MSG format requires Outlook COM objects. This module exports to EML which can be opened in Outlook.
#>

# Import required modules

<#
.SYNOPSIS
    Exports a single email to EML format (MIME)
.PARAMETER UserId
    User email address
.PARAMETER MessageId
    Message ID
.PARAMETER SavePath
    Path to save the EML file
.PARAMETER FileName
    Optional custom filename
#>
function Export-EmailToEML {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$MessageId,

        [Parameter(Mandatory = $true)]
        [string]$SavePath,

        [Parameter(Mandatory = $false)]
        [string]$FileName
    )

    try {
        # Get message details
        $message = Get-MgUserMessage -UserId $UserId -MessageId $MessageId -ErrorAction Stop

        # Get MIME content using Graph API
        Write-Verbose "Retrieving MIME content for message $MessageId"
        $mimeUri = "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId/`$value"
        $mimeContent = Invoke-MgGraphRequest -Method GET -Uri $mimeUri -ErrorAction Stop

        # Generate filename if not provided
        if ([string]::IsNullOrWhiteSpace($FileName)) {
            $dateStr = $message.ReceivedDateTime.ToString("yy-MM-dd")
            $senderPart = ($message.From.EmailAddress.Address -split '@')[0]
            $senderPart = Get-SafeFilename -Text $senderPart -MaxLength 30
            $subjectPart = Get-SafeFilename -Text $message.Subject -MaxLength 50
            $FileName = "$dateStr $senderPart - $subjectPart.eml"
        }

        # Ensure .eml extension
        if (-not $FileName.EndsWith('.eml')) {
            $FileName += '.eml'
        }

        # Ensure directory exists
        if (-not (Ensure-DirectoryExists -Path $SavePath)) {
            return $null
        }

        # Get unique file path
        $fullPath = Join-Path -Path $SavePath -ChildPath $FileName
        $fullPath = Get-UniqueFilePath -FilePath $fullPath

        # Save MIME content
        Set-Content -Path $fullPath -Value $mimeContent -Encoding UTF8 -ErrorAction Stop

        Write-Host "Exported: $FileName" -ForegroundColor Green

        return @{
            Success = $true
            FilePath = $fullPath
            FileName = $FileName
        }

    } catch {
        Write-Error "Error exporting email to EML: $($_.Exception.Message)"
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

<#
.SYNOPSIS
    Exports a single email to MSG format using Outlook COM (if available)
.PARAMETER UserId
    User email address
.PARAMETER MessageId
    Message ID
.PARAMETER SavePath
    Path to save the MSG file
.PARAMETER FileName
    Optional custom filename
#>
function Export-EmailToMSG {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$MessageId,

        [Parameter(Mandatory = $true)]
        [string]$SavePath,

        [Parameter(Mandatory = $false)]
        [string]$FileName
    )

    # Check if Outlook is available
    try {
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $outlookAvailable = $true
    } catch {
        Write-Warning "Outlook COM object not available. Exporting to EML instead."
        $outlookAvailable = $false
    }

    if (-not $outlookAvailable) {
        # Fallback to EML export
        if ($FileName -and $FileName.EndsWith('.msg')) {
            $FileName = $FileName -replace '\.msg$', '.eml'
        }
        return Export-EmailToEML -UserId $UserId -MessageId $MessageId -SavePath $SavePath -FileName $FileName
    }

    try {
        # Get message details
        $message = Get-MgUserMessage -UserId $UserId -MessageId $MessageId -ErrorAction Stop

        # First export to EML
        $tempEmlPath = Join-Path -Path $env:TEMP -ChildPath "$MessageId.eml"
        $mimeUri = "https://graph.microsoft.com/v1.0/users/$UserId/messages/$MessageId/`$value"
        $mimeContent = Invoke-MgGraphRequest -Method GET -Uri $mimeUri -ErrorAction Stop
        Set-Content -Path $tempEmlPath -Value $mimeContent -Encoding UTF8 -ErrorAction Stop

        # Import EML into Outlook and save as MSG
        $outlookNamespace = $outlook.GetNamespace("MAPI")
        $mailItem = $outlookNamespace.OpenSharedItem($tempEmlPath)

        # Generate filename if not provided
        if ([string]::IsNullOrWhiteSpace($FileName)) {
            $dateStr = $message.ReceivedDateTime.ToString("yy-MM-dd")
            $senderPart = ($message.From.EmailAddress.Address -split '@')[0]
            $senderPart = Get-SafeFilename -Text $senderPart -MaxLength 30
            $subjectPart = Get-SafeFilename -Text $message.Subject -MaxLength 50
            $FileName = "$dateStr $senderPart - $subjectPart.msg"
        }

        # Ensure .msg extension
        if (-not $FileName.EndsWith('.msg')) {
            $FileName += '.msg'
        }

        # Ensure directory exists
        if (-not (Ensure-DirectoryExists -Path $SavePath)) {
            return $null
        }

        # Get unique file path
        $fullPath = Join-Path -Path $SavePath -ChildPath $FileName
        $fullPath = Get-UniqueFilePath -FilePath $fullPath

        # Save as MSG
        $mailItem.SaveAs($fullPath, 3) # 3 = olMSG format

        # Cleanup
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mailItem) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlookNamespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        Remove-Item -Path $tempEmlPath -Force -ErrorAction SilentlyContinue

        Write-Host "Exported: $FileName" -ForegroundColor Green

        return @{
            Success = $true
            FilePath = $fullPath
            FileName = $FileName
        }

    } catch {
        Write-Error "Error exporting email to MSG: $($_.Exception.Message)"
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

<#
.SYNOPSIS
    Bulk exports emails to EML/MSG format with time filters
.PARAMETER UserId
    User email address
.PARAMETER TimeFilter
    Time filter: 'LastDay', 'Last7Days', 'Last30Days', 'Last90Days', 'LastWeek', 'LastMonth', 'All'
.PARAMETER SavePath
    Path to save emails
.PARAMETER Format
    Export format: 'EML' or 'MSG'
.PARAMETER SenderDomain
    Optional: Filter by sender domain
#>
function Export-BulkEmails {
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
        [ValidateSet('EML', 'MSG')]
        [string]$Format = 'EML',

        [Parameter(Mandatory = $false)]
        [string]$SenderDomain
    )

    Write-Host "`n=== Bulk Email Export ===" -ForegroundColor Cyan
    Write-Host "Time Filter: $TimeFilter" -ForegroundColor Yellow
    Write-Host "Format: $Format" -ForegroundColor Yellow
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

    $odataFilter = if ($filters.Count -gt 0) { $filters -join ' and ' } else { $null }

    try {
        # Import GraphApiService
        Import-Module (Join-Path $PSScriptRoot "..\Core\GraphApiService.psm1") -Force

        Write-Host "`nFetching messages..." -ForegroundColor Cyan

        $params = @{
            UserId = $UserId
            All = $true
        }
        if ($odataFilter) {
            $params.Filter = $odataFilter
        }

        $messages = Get-GraphMessages @params

        if ($null -eq $messages -or $messages.Count -eq 0) {
            Write-Warning "No messages found matching criteria"
            return
        }

        Write-Host "Found $($messages.Count) messages" -ForegroundColor Green

        $totalExported = 0
        $totalErrors = 0
        $processedMessages = 0

        foreach ($message in $messages) {
            $processedMessages++
            Write-Progress -Activity "Exporting Emails" `
                -Status "Processing message $processedMessages of $($messages.Count)" `
                -PercentComplete (($processedMessages / $messages.Count) * 100)

            try {
                if ($Format -eq 'MSG') {
                    $result = Export-EmailToMSG -UserId $UserId -MessageId $message.Id -SavePath $SavePath
                } else {
                    $result = Export-EmailToEML -UserId $UserId -MessageId $message.Id -SavePath $SavePath
                }

                if ($result -and $result.Success) {
                    $totalExported++
                } else {
                    $totalErrors++
                }

            } catch {
                Write-Warning "Error exporting message: $($_.Exception.Message)"
                $totalErrors++
            }
        }

        Write-Progress -Activity "Exporting Emails" -Completed

        # Summary
        Write-Host "`n=== Export Summary ===" -ForegroundColor Cyan
        Write-Host "Total Exported: $totalExported" -ForegroundColor Green
        Write-Host "Total Errors: $totalErrors" -ForegroundColor Red
        Write-Host "Save Location: $SavePath" -ForegroundColor Cyan

    } catch {
        Write-Error "Error during bulk export: $($_.Exception.Message)"
        Write-Progress -Activity "Exporting Emails" -Completed
    }
}

Export-ModuleMember -Function Export-EmailToEML, Export-EmailToMSG, Export-BulkEmails
