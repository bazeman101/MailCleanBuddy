<#
.SYNOPSIS
    VIP Manager module for MailCleanBuddy
.DESCRIPTION
    Manages VIP (Very Important Person) senders with whitelist protection.
    Prevents accidental deletion of important emails and provides special handling.
#>

# Import dependencies

# VIP database file path
$script:VIPDatabasePath = $null

# Function: Initialize-VIPDatabase
function Initialize-VIPDatabase {
    <#
    .SYNOPSIS
        Initializes VIP database
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
    $script:VIPDatabasePath = Join-Path $PSScriptRoot "..\..\vip_senders_$sanitizedEmail.json"

    if (-not (Test-Path $script:VIPDatabasePath)) {
        $initialData = @{
            VIPs = @()
            LastUpdated = (Get-Date).ToString("o")
        }
        $initialData | ConvertTo-Json -Depth 10 | Set-Content -Path $script:VIPDatabasePath -Encoding UTF8
    }
}

# Function: Get-VIPList
function Get-VIPList {
    <#
    .SYNOPSIS
        Gets list of VIP senders
    .OUTPUTS
        Array of VIP objects
    #>
    [CmdletBinding()]
    param()

    try {
        if (-not $script:VIPDatabasePath -or -not (Test-Path $script:VIPDatabasePath)) {
            return @()
        }

        $data = Get-Content -Path $script:VIPDatabasePath -Raw | ConvertFrom-Json
        return $data.VIPs
    }
    catch {
        Write-Warning "Error loading VIP list: $($_.Exception.Message)"
        return @()
    }
}

# Function: Add-VIPSender
function Add-VIPSender {
    <#
    .SYNOPSIS
        Adds sender to VIP list
    .PARAMETER EmailAddress
        Email address to add
    .PARAMETER Name
        Sender name (optional)
    .PARAMETER Notes
        Notes about this VIP (optional)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,

        [Parameter(Mandatory = $false)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string]$Notes
    )

    try {
        $data = Get-Content -Path $script:VIPDatabasePath -Raw | ConvertFrom-Json

        # Check if already exists
        $existing = $data.VIPs | Where-Object { $_.EmailAddress -eq $EmailAddress }
        if ($existing) {
            Write-Warning (Get-LocalizedString "vip_alreadyExists" -FormatArgs @($EmailAddress))
            return $false
        }

        # Add new VIP
        $newVIP = [PSCustomObject]@{
            EmailAddress = $EmailAddress.ToLower()
            Name = if ($Name) { $Name } else { $EmailAddress }
            Notes = $Notes
            AddedDate = (Get-Date).ToString("o")
            EmailCount = 0
        }

        $data.VIPs += $newVIP
        $data.LastUpdated = (Get-Date).ToString("o")

        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:VIPDatabasePath -Encoding UTF8

        Write-Host (Get-LocalizedString "vip_added" -FormatArgs @($EmailAddress)) -ForegroundColor $Global:ColorScheme.Success
        return $true
    }
    catch {
        Write-Error "Error adding VIP: $($_.Exception.Message)"
        return $false
    }
}

# Function: Remove-VIPSender
function Remove-VIPSender {
    <#
    .SYNOPSIS
        Removes sender from VIP list
    .PARAMETER EmailAddress
        Email address to remove
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )

    try {
        $data = Get-Content -Path $script:VIPDatabasePath -Raw | ConvertFrom-Json

        $initialCount = $data.VIPs.Count
        $data.VIPs = @($data.VIPs | Where-Object { $_.EmailAddress -ne $EmailAddress.ToLower() })

        if ($data.VIPs.Count -eq $initialCount) {
            Write-Warning (Get-LocalizedString "vip_notFound" -FormatArgs @($EmailAddress))
            return $false
        }

        $data.LastUpdated = (Get-Date).ToString("o")
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:VIPDatabasePath -Encoding UTF8

        Write-Host (Get-LocalizedString "vip_removed" -FormatArgs @($EmailAddress)) -ForegroundColor $Global:ColorScheme.Success
        return $true
    }
    catch {
        Write-Error "Error removing VIP: $($_.Exception.Message)"
        return $false
    }
}

# Function: Test-IsVIPSender
function Test-IsVIPSender {
    <#
    .SYNOPSIS
        Checks if sender is a VIP
    .PARAMETER EmailAddress
        Email address to check
    .OUTPUTS
        Boolean
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )

    $vips = Get-VIPList
    return ($vips | Where-Object { $_.EmailAddress -eq $EmailAddress.ToLower() }).Count -gt 0
}

# Function: Update-VIPStatistics
function Update-VIPStatistics {
    <#
    .SYNOPSIS
        Updates VIP statistics from cache
    #>
    [CmdletBinding()]
    param()

    try {
        $vips = Get-VIPList
        if ($vips.Count -eq 0) {
            return
        }

        $cache = Get-SenderCache
        if (-not $cache) {
            return
        }

        $data = Get-Content -Path $script:VIPDatabasePath -Raw | ConvertFrom-Json

        foreach ($vip in $data.VIPs) {
            $emailCount = 0

            foreach ($domain in $cache.Keys) {
                $domainEmails = $cache[$domain].Messages | Where-Object {
                    $_.SenderEmailAddress -eq $vip.EmailAddress
                }
                $emailCount += $domainEmails.Count
            }

            $vip.EmailCount = $emailCount
        }

        $data.LastUpdated = (Get-Date).ToString("o")
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:VIPDatabasePath -Encoding UTF8
    }
    catch {
        Write-Warning "Error updating VIP statistics: $($_.Exception.Message)"
    }
}

# Function: Show-VIPManager
function Show-VIPManager {
    <#
    .SYNOPSIS
        Interactive VIP manager interface
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        # Initialize database
        Initialize-VIPDatabase -UserEmail $UserEmail

        while ($true) {
            Clear-Host

            # Display header
            $title = Get-LocalizedString "vip_title" -FormatArgs @($UserEmail)
            Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
            Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""

            Write-Host (Get-LocalizedString "vip_description") -ForegroundColor $Global:ColorScheme.Info
            Write-Host ""

            # Update statistics
            Update-VIPStatistics

            # Get VIP list
            $vips = Get-VIPList

            if ($vips.Count -eq 0) {
                Write-Host (Get-LocalizedString "vip_noVIPs") -ForegroundColor $Global:ColorScheme.Warning
            } else {
                Write-Host (Get-LocalizedString "vip_listTitle" -FormatArgs @($vips.Count)) -ForegroundColor $Global:ColorScheme.SectionHeader
                Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

                $format = "{0,-4} {1,-35} {2,-35} {3,10} {4,15}"
                Write-Host ($format -f "#", "Name", "Email", "Emails", "Added") -ForegroundColor $Global:ColorScheme.Header
                Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

                $index = 1
                foreach ($vip in $vips) {
                    $name = if ($vip.Name.Length -gt 34) {
                        $vip.Name.Substring(0, 31) + "..."
                    } else {
                        $vip.Name
                    }

                    $email = if ($vip.EmailAddress.Length -gt 34) {
                        $vip.EmailAddress.Substring(0, 31) + "..."
                    } else {
                        $vip.EmailAddress
                    }

                    $addedDate = ConvertTo-SafeDateTime -DateTimeValue $vip.AddedDate.ToString("yyyy-MM-dd")

                    Write-Host ($format -f $index, $name, $email, $vip.EmailCount, $addedDate) -ForegroundColor $Global:ColorScheme.Value
                    $index++
                }
            }

            Write-Host ""

            # Menu
            Write-Host (Get-LocalizedString "vip_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  1. $(Get-LocalizedString 'vip_addVIP')" -ForegroundColor Green
            Write-Host "  2. $(Get-LocalizedString 'vip_removeVIP')" -ForegroundColor Yellow
            Write-Host "  3. $(Get-LocalizedString 'vip_addFromCache')" -ForegroundColor Cyan
            Write-Host "  4. $(Get-LocalizedString 'vip_exportList')" -ForegroundColor Magenta
            Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
            Write-Host ""

            $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

            switch ($choice.ToUpper()) {
                "1" {
                    # Add VIP manually
                    Write-Host ""
                    $email = Read-Host (Get-LocalizedString "vip_enterEmail")
                    if (-not [string]::IsNullOrWhiteSpace($email)) {
                        $name = Read-Host (Get-LocalizedString "vip_enterName")
                        $notes = Read-Host (Get-LocalizedString "vip_enterNotes")
                        Add-VIPSender -EmailAddress $email -Name $name -Notes $notes
                    }
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "2" {
                    # Remove VIP
                    if ($vips.Count -eq 0) {
                        Write-Host (Get-LocalizedString "vip_noVIPs") -ForegroundColor $Global:ColorScheme.Warning
                    } else {
                        Write-Host ""
                        $vipNum = Read-Host (Get-LocalizedString "vip_enterNumber")
                        if ($vipNum -match '^\d+$' -and [int]$vipNum -ge 1 -and [int]$vipNum -le $vips.Count) {
                            $selectedVIP = $vips[[int]$vipNum - 1]
                            $confirm = Show-Confirmation -Message (Get-LocalizedString "vip_confirmRemove" -FormatArgs @($selectedVIP.EmailAddress))
                            if ($confirm) {
                                Remove-VIPSender -EmailAddress $selectedVIP.EmailAddress
                            }
                        } else {
                            Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
                        }
                    }
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "3" {
                    # Add from cache
                    Show-AddVIPFromCache
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "4" {
                    # Export list
                    Export-VIPList -VIPs $vips
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "Q" {
                    return
                }
                default {
                    Write-Host (Get-LocalizedString "unsubscribe_invalidChoice") -ForegroundColor $Global:ColorScheme.Warning
                    Start-Sleep -Seconds 1
                }
            }
        }
    }
    catch {
        Write-Error "Error in VIP manager: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Show-AddVIPFromCache
function Show-AddVIPFromCache {
    <#
    .SYNOPSIS
        Shows top senders to add as VIP
    #>
    [CmdletBinding()]
    param()

    $cache = Get-SenderCache
    if (-not $cache -or $cache.Count -eq 0) {
        Write-Host (Get-LocalizedString "analytics_noCacheData") -ForegroundColor $Global:ColorScheme.Warning
        return
    }

    # Get top senders
    $topSenders = @()
    foreach ($domain in $cache.Keys) {
        $senderData = $cache[$domain]
        if ($senderData.Messages.Count -gt 0) {
            $sampleMsg = $senderData.Messages[0]
            $topSenders += [PSCustomObject]@{
                Domain = $domain
                EmailAddress = $sampleMsg.SenderEmailAddress
                Name = $sampleMsg.SenderName
                Count = $senderData.Count
            }
        }
    }

    $topSenders = $topSenders | Sort-Object -Property Count -Descending | Select-Object -First 20

    Write-Host ""
    Write-Host (Get-LocalizedString "vip_topSenders") -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

    $format = "{0,-4} {1,-40} {2,10}"
    Write-Host ($format -f "#", "Sender", "Count") -ForegroundColor $Global:ColorScheme.Header
    Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

    $index = 1
    foreach ($sender in $topSenders) {
        $display = "$($sender.Name) <$($sender.EmailAddress)>"
        if ($display.Length -gt 39) {
            $display = $display.Substring(0, 36) + "..."
        }

        Write-Host ($format -f $index, $display, $sender.Count) -ForegroundColor $Global:ColorScheme.Normal
        $index++
    }

    Write-Host ""
    $senderNum = Read-Host (Get-LocalizedString "vip_selectSender")

    if ($senderNum -match '^\d+$' -and [int]$senderNum -ge 1 -and [int]$senderNum -le $topSenders.Count) {
        $selected = $topSenders[[int]$senderNum - 1]
        $notes = Read-Host (Get-LocalizedString "vip_enterNotes")
        Add-VIPSender -EmailAddress $selected.EmailAddress -Name $selected.Name -Notes $notes
    } else {
        Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
    }
}

# Function: Export-VIPList
function Export-VIPList {
    <#
    .SYNOPSIS
        Exports VIP list to CSV
    .PARAMETER VIPs
        VIP list
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$VIPs
    )

    if ($VIPs.Count -eq 0) {
        Write-Host (Get-LocalizedString "vip_noVIPs") -ForegroundColor $Global:ColorScheme.Warning
        return
    }

    $defaultPath = Join-Path $PSScriptRoot "..\..\vip_list.csv"
    $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    $VIPs | Select-Object Name, EmailAddress, EmailCount, Notes, AddedDate |
        Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
}

# Export functions
Export-ModuleMember -Function Show-VIPManager, Add-VIPSender, Remove-VIPSender, `
    Test-IsVIPSender, Get-VIPList, Initialize-VIPDatabase, Update-VIPStatistics
