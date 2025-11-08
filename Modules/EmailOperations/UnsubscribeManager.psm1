<#
.SYNOPSIS
    Unsubscribe Manager module for MailCleanBuddy
.DESCRIPTION
    Provides functionality to identify and manage newsletter/marketing email subscriptions
    including unsubscribe link detection and bulk management of unwanted senders.
#>

# Import dependencies

# Function: Show-UnsubscribeOpportunities
function Show-UnsubscribeOpportunities {
    <#
    .SYNOPSIS
        Shows interactive list of potential newsletter/marketing senders
    .DESCRIPTION
        Displays identified newsletter senders with options to manage them
    .PARAMETER UserEmail
        Email address of the mailbox
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Clear-Host

        # Display header
        $title = Get-LocalizedString "unsubscribe_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host (Get-LocalizedString "unsubscribe_scanning") -ForegroundColor $Global:ColorScheme.Info

        # Import analytics function
        Import-Module (Join-Path $PSScriptRoot "..\Analytics\AnalyticsDashboard.psm1") -Force
        $opportunities = Get-UnsubscribeOpportunities -MinEmailCount 3

        if ($opportunities.Count -eq 0) {
            Write-Host "`n$(Get-LocalizedString 'unsubscribe_noOpportunities')" -ForegroundColor $Global:ColorScheme.Success
            Write-Host ""
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        Write-Host "`n$(Get-LocalizedString 'unsubscribe_foundCount' -FormatArgs @($opportunities.Count))" -ForegroundColor $Global:ColorScheme.Success
        Write-Host ""

        # Display opportunities in a table
        $format = "{0,-4} {1,-45} {2,8} {3,12} {4,10}"
        Write-Host ($format -f "#", (Get-LocalizedString 'senderOverview_headerDomain'),
            (Get-LocalizedString 'senderOverview_headerCount'), "Storage (MB)", "Score") -ForegroundColor $Global:ColorScheme.Header
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        $index = 1
        foreach ($opp in $opportunities) {
            $domainDisplay = if ($opp.Domain.Length -gt 44) {
                $opp.Domain.Substring(0, 41) + "..."
            } else {
                $opp.Domain
            }

            Write-Host ($format -f $index, $domainDisplay, $opp.EmailCount, $opp.TotalSizeMB, $opp.Score) -ForegroundColor $Global:ColorScheme.Normal
            $index++
        }

        Write-Host "`n$(Get-LocalizedString 'unsubscribe_instructions')" -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Menu for actions
        while ($true) {
            Write-Host (Get-LocalizedString "unsubscribe_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  1. $(Get-LocalizedString 'unsubscribe_viewDetails')" -ForegroundColor Green
            Write-Host "  2. $(Get-LocalizedString 'unsubscribe_deleteFromDomain')" -ForegroundColor Yellow
            Write-Host "  3. $(Get-LocalizedString 'unsubscribe_moveToFolder')" -ForegroundColor Cyan
            Write-Host "  4. $(Get-LocalizedString 'unsubscribe_exportList')" -ForegroundColor Magenta
            Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
            Write-Host ""

            $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

            switch ($choice.ToUpper()) {
                "1" {
                    # View details for a specific sender
                    $domainNum = Read-Host (Get-LocalizedString "unsubscribe_enterDomainNumber")
                    if ($domainNum -match '^\d+$' -and [int]$domainNum -ge 1 -and [int]$domainNum -le $opportunities.Count) {
                        $selected = $opportunities[[int]$domainNum - 1]
                        Show-UnsubscribeDetails -Domain $selected.Domain -Opportunity $selected
                    } else {
                        Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
                    }
                }
                "2" {
                    # Delete all emails from a domain
                    $domainNum = Read-Host (Get-LocalizedString "unsubscribe_enterDomainNumber")
                    if ($domainNum -match '^\d+$' -and [int]$domainNum -ge 1 -and [int]$domainNum -le $opportunities.Count) {
                        $selected = $opportunities[[int]$domainNum - 1]
                        Remove-NewsletterEmails -UserEmail $UserEmail -Domain $selected.Domain -EmailCount $selected.EmailCount
                    } else {
                        Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
                    }
                }
                "3" {
                    # Move all emails to a folder
                    $domainNum = Read-Host (Get-LocalizedString "unsubscribe_enterDomainNumber")
                    if ($domainNum -match '^\d+$' -and [int]$domainNum -ge 1 -and [int]$domainNum -le $opportunities.Count) {
                        $selected = $opportunities[[int]$domainNum - 1]
                        Move-NewsletterEmails -UserEmail $UserEmail -Domain $selected.Domain -EmailCount $selected.EmailCount
                    } else {
                        Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
                    }
                }
                "4" {
                    # Export list
                    Export-UnsubscribeList -Opportunities $opportunities
                }
                "Q" {
                    return
                }
                default {
                    Write-Host (Get-LocalizedString "unsubscribe_invalidChoice") -ForegroundColor $Global:ColorScheme.Warning
                }
            }

            Write-Host ""
        }
    }
    catch {
        Write-Error "Error showing unsubscribe opportunities: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Show-UnsubscribeDetails
function Show-UnsubscribeDetails {
    <#
    .SYNOPSIS
        Shows detailed information about a newsletter sender
    .PARAMETER Domain
        The sender domain
    .PARAMETER Opportunity
        The opportunity object with details
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Opportunity
    )

    Clear-Host
    Write-Host "`n$(Get-LocalizedString 'unsubscribe_detailsTitle' -FormatArgs @($Domain))" -ForegroundColor $Global:ColorScheme.Highlight
    Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
    Write-Host ""

    Write-Host "  $(Get-LocalizedString 'senderOverview_headerCount'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $Opportunity.EmailCount -ForegroundColor $Global:ColorScheme.Value

    Write-Host "  Storage (MB): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $Opportunity.TotalSizeMB -ForegroundColor $Global:ColorScheme.Value

    Write-Host "  $(Get-LocalizedString 'unsubscribe_detectionScore'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $Opportunity.Score -ForegroundColor $Global:ColorScheme.Value

    Write-Host "  $(Get-LocalizedString 'unsubscribe_reasons'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
    Write-Host $Opportunity.Reasons -ForegroundColor $Global:ColorScheme.Muted

    if ($Opportunity.SampleSender) {
        Write-Host "  $(Get-LocalizedString 'unsubscribe_sampleSender'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host $Opportunity.SampleSender -ForegroundColor $Global:ColorScheme.Value
    }

    if ($Opportunity.SampleSubject) {
        Write-Host "  $(Get-LocalizedString 'unsubscribe_sampleSubject'): " -NoNewline -ForegroundColor $Global:ColorScheme.Label
        Write-Host $Opportunity.SampleSubject -ForegroundColor $Global:ColorScheme.Value
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_manualUnsubscribe") -ForegroundColor $Global:ColorScheme.Info
    Write-Host (Get-LocalizedString "unsubscribe_useMenuOptions") -ForegroundColor $Global:ColorScheme.Info
    Write-Host ""

    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Function: Remove-NewsletterEmails
function Remove-NewsletterEmails {
    <#
    .SYNOPSIS
        Deletes all emails from a newsletter sender
    .PARAMETER UserEmail
        User email address
    .PARAMETER Domain
        Sender domain
    .PARAMETER EmailCount
        Number of emails to delete
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$Domain,

        [Parameter(Mandatory = $true)]
        [int]$EmailCount
    )

    Write-Host ""
    $confirm = Show-Confirmation -Message (Get-LocalizedString "confirmation_promptDeleteAllDomainEmails" -FormatArgs @($EmailCount, $Domain))

    if (-not $confirm) {
        Write-Host (Get-LocalizedString "performActionAll_deleteCancelled") -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "performActionAll_startingDelete" -FormatArgs @($EmailCount, $Domain)) -ForegroundColor $Global:ColorScheme.Info

    # Get emails from cache
    $cache = Get-SenderCache
    if (-not $cache.ContainsKey($Domain)) {
        Write-Host (Get-LocalizedString "manageEmails_domainNotFoundInCache" -FormatArgs @($Domain)) -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    $messages = $cache[$Domain].Messages
    $successCount = 0
    $errorCount = 0

    foreach ($msg in $messages) {
        try {
            Remove-GraphMessage -UserId $UserEmail -MessageId $msg.MessageId | Out-Null
            $successCount++

            Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityDelete") `
                          -Status (Get-LocalizedString "performActionAll_progressStatusDelete" -FormatArgs @($msg.Subject)) `
                          -PercentComplete (($successCount / $messages.Count) * 100)
        }
        catch {
            Write-Warning (Get-LocalizedString "performActionAll_errorDeletingEmailId" -FormatArgs @($msg.MessageId, $_.Exception.Message))
            $errorCount++
        }
    }

    Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityDelete") -Completed

    # Update cache
    Update-CacheAfterAction -Domain $Domain -MessageIds @() -IsDeleteAll $true

    Write-Host ""
    Write-Host (Get-LocalizedString "performActionAll_deleteComplete" -FormatArgs @($successCount)) -ForegroundColor $Global:ColorScheme.Success
    if ($errorCount -gt 0) {
        Write-Host (Get-LocalizedString "performActionAll_deleteErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
    }

    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Function: Move-NewsletterEmails
function Move-NewsletterEmails {
    <#
    .SYNOPSIS
        Moves all emails from a newsletter sender to a folder
    .PARAMETER UserEmail
        User email address
    .PARAMETER Domain
        Sender domain
    .PARAMETER EmailCount
        Number of emails to move
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [string]$Domain,

        [Parameter(Mandatory = $true)]
        [int]$EmailCount
    )

    # Select destination folder
    $destinationFolder = Select-MailFolder -UserId $UserEmail

    if (-not $destinationFolder) {
        Write-Host (Get-LocalizedString "performActionAll_moveNoDestination") -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    Write-Host ""
    $confirm = Show-Confirmation -Message (Get-LocalizedString "confirmation_promptMoveAllDomainEmails" -FormatArgs @($EmailCount, $Domain, $destinationFolder.displayName))

    if (-not $confirm) {
        Write-Host (Get-LocalizedString "performActionAll_moveCancelled") -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "performActionAll_startingMove" -FormatArgs @($EmailCount, $Domain, $destinationFolder.displayName)) -ForegroundColor $Global:ColorScheme.Info

    # Get emails from cache
    $cache = Get-SenderCache
    if (-not $cache.ContainsKey($Domain)) {
        Write-Host (Get-LocalizedString "manageEmails_domainNotFoundInCache" -FormatArgs @($Domain)) -ForegroundColor $Global:ColorScheme.Warning
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
        return
    }

    $messages = $cache[$Domain].Messages
    $successCount = 0
    $errorCount = 0

    foreach ($msg in $messages) {
        try {
            Move-GraphMessage -UserId $UserEmail -MessageId $msg.MessageId -DestinationFolderId $destinationFolder.id | Out-Null
            $successCount++

            Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityMove") `
                          -Status (Get-LocalizedString "performActionAll_progressStatusMove" -FormatArgs @($msg.Subject)) `
                          -PercentComplete (($successCount / $messages.Count) * 100)
        }
        catch {
            Write-Warning (Get-LocalizedString "performActionAll_errorMovingEmailId" -FormatArgs @($msg.MessageId, $_.Exception.Message))
            $errorCount++
        }
    }

    Write-Progress -Activity (Get-LocalizedString "performActionAll_progressActivityMove") -Completed

    # Update cache
    Update-CacheAfterAction -Domain $Domain -MessageIds @() -IsDeleteAll $true

    Write-Host ""
    Write-Host (Get-LocalizedString "performActionAll_moveComplete" -FormatArgs @($successCount)) -ForegroundColor $Global:ColorScheme.Success
    if ($errorCount -gt 0) {
        Write-Host (Get-LocalizedString "performActionAll_moveErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
    }

    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
}

# Function: Export-UnsubscribeList
function Export-UnsubscribeList {
    <#
    .SYNOPSIS
        Exports the list of newsletter opportunities to a CSV file
    .PARAMETER Opportunities
        Array of opportunity objects
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Opportunities
    )

    try {
        $defaultPath = Join-Path $PSScriptRoot "..\..\unsubscribe_opportunities.csv"
        $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

        if ([string]::IsNullOrWhiteSpace($exportPath)) {
            $exportPath = $defaultPath
        }

        # Export to CSV
        $Opportunities | Select-Object Domain, EmailCount, TotalSizeMB, Score, Reasons, SampleSender, SampleSubject |
            Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

        Write-Host ""
        Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
    catch {
        Write-Error "Error exporting unsubscribe list: $($_.Exception.Message)"
        Write-Host (Get-LocalizedString "unsubscribe_exportError" -FormatArgs @($_.Exception.Message)) -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Export functions
Export-ModuleMember -Function Show-UnsubscribeOpportunities, Show-UnsubscribeDetails, `
    Remove-NewsletterEmails, Move-NewsletterEmails, Export-UnsubscribeList
