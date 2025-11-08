<#
.SYNOPSIS
    Smart Folder Auto-Organizer module for MailCleanBuddy
.DESCRIPTION
    Learns from user actions to suggest automatic email organization rules.
    Tracks patterns in move/delete operations and generates folder organization suggestions.
#>

# Import dependencies

# Smart rules database path
$script:SmartRulesPath = $null

# Function: Initialize-SmartOrganizer
function Initialize-SmartOrganizer {
    <#
    .SYNOPSIS
        Initializes smart organizer database
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
    $script:SmartRulesPath = Join-Path $PSScriptRoot "..\..\smart_rules_$sanitizedEmail.json"

    if (-not (Test-Path $script:SmartRulesPath)) {
        $initialData = @{
            Actions = @()
            SuggestedRules = @()
            AppliedRules = @()
            LastAnalysis = $null
        }
        $initialData | ConvertTo-Json -Depth 10 | Set-Content -Path $script:SmartRulesPath -Encoding UTF8
    }
}

# Function: Track-UserAction
function Track-UserAction {
    <#
    .SYNOPSIS
        Tracks a user action for learning
    .PARAMETER ActionType
        Type: 'Move', 'Delete', 'Archive'
    .PARAMETER SenderDomain
        Sender domain
    .PARAMETER DestinationFolder
        Destination folder (if move)
    .PARAMETER EmailCount
        Number of emails affected
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Move', 'Delete', 'Archive')]
        [string]$ActionType,

        [Parameter(Mandatory = $true)]
        [string]$SenderDomain,

        [Parameter(Mandatory = $false)]
        [string]$DestinationFolder,

        [Parameter(Mandatory = $false)]
        [int]$EmailCount = 1
    )

    try {
        if (-not $script:SmartRulesPath) {
            return
        }

        $data = Get-Content -Path $script:SmartRulesPath -Raw | ConvertFrom-Json

        $action = [PSCustomObject]@{
            ActionType = $ActionType
            SenderDomain = $SenderDomain.ToLower()
            DestinationFolder = $DestinationFolder
            EmailCount = $EmailCount
            Timestamp = (Get-Date).ToString("o")
        }

        $data.Actions += $action
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:SmartRulesPath -Encoding UTF8
    }
    catch {
        Write-Warning "Error tracking action: $($_.Exception.Message)"
    }
}

# Function: Analyze-UserPatterns
function Analyze-UserPatterns {
    <#
    .SYNOPSIS
        Analyzes user actions to generate rule suggestions
    .OUTPUTS
        Array of suggested rules
    #>
    [CmdletBinding()]
    param()

    try {
        $data = Get-Content -Path $script:SmartRulesPath -Raw | ConvertFrom-Json

        if ($data.Actions.Count -lt 3) {
            return @()  # Need at least 3 actions to learn
        }

        # Group actions by sender domain
        $grouped = $data.Actions | Group-Object -Property SenderDomain

        $suggestions = @()

        foreach ($group in $grouped) {
            $domain = $group.Name
            $actions = $group.Group

            # Check for consistent move pattern
            $moves = $actions | Where-Object { $_.ActionType -eq 'Move' }
            if ($moves.Count -ge 2) {
                # Find most common destination
                $destinations = $moves | Group-Object -Property DestinationFolder | Sort-Object Count -Descending
                $topDestination = $destinations[0]

                if ($topDestination.Count -ge 2) {
                    $confidence = [math]::Round(($topDestination.Count / $moves.Count) * 100, 0)

                    $suggestions += [PSCustomObject]@{
                        Type = "Move"
                        SenderDomain = $domain
                        DestinationFolder = $topDestination.Name
                        Confidence = $confidence
                        BasedOnActions = $topDestination.Count
                        Description = "Move emails from '$domain' to '$($topDestination.Name)'"
                    }
                }
            }

            # Check for consistent delete pattern
            $deletes = $actions | Where-Object { $_.ActionType -eq 'Delete' }
            if ($deletes.Count -ge 3) {
                $confidence = [math]::Round(($deletes.Count / $actions.Count) * 100, 0)

                if ($confidence -ge 70) {
                    $suggestions += [PSCustomObject]@{
                        Type = "Delete"
                        SenderDomain = $domain
                        DestinationFolder = $null
                        Confidence = $confidence
                        BasedOnActions = $deletes.Count
                        Description = "Auto-delete emails from '$domain'"
                    }
                }
            }

            # Check for consistent archive pattern
            $archives = $actions | Where-Object { $_.ActionType -eq 'Archive' }
            if ($archives.Count -ge 2) {
                $confidence = [math]::Round(($archives.Count / $actions.Count) * 100, 0)

                if ($confidence -ge 60) {
                    $suggestions += [PSCustomObject]@{
                        Type = "Archive"
                        SenderDomain = $domain
                        DestinationFolder = "Archive"
                        Confidence = $confidence
                        BasedOnActions = $archives.Count
                        Description = "Auto-archive emails from '$domain'"
                    }
                }
            }
        }

        # Save suggestions
        $data.SuggestedRules = $suggestions
        $data.LastAnalysis = (Get-Date).ToString("o")
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:SmartRulesPath -Encoding UTF8

        return $suggestions | Sort-Object -Property Confidence -Descending
    }
    catch {
        Write-Error "Error analyzing patterns: $($_.Exception.Message)"
        return @()
    }
}

# Function: Show-SmartOrganizer
function Show-SmartOrganizer {
    <#
    .SYNOPSIS
        Interactive smart organizer interface
    .PARAMETER UserEmail
        User email address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Initialize-SmartOrganizer -UserEmail $UserEmail

        Clear-Host

        $title = Get-LocalizedString "smart_title" -FormatArgs @($UserEmail)
        Write-Host "`n$title" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        Write-Host (Get-LocalizedString "smart_description") -ForegroundColor $Global:ColorScheme.Info
        Write-Host ""

        # Analyze patterns
        Write-Host (Get-LocalizedString "smart_analyzing") -ForegroundColor $Global:ColorScheme.Info
        $suggestions = Analyze-UserPatterns

        if ($suggestions.Count -eq 0) {
            Write-Host "`n$(Get-LocalizedString 'smart_noSuggestions')" -ForegroundColor $Global:ColorScheme.Warning
            Write-Host (Get-LocalizedString "smart_needMoreActions") -ForegroundColor $Global:ColorScheme.Info
            Write-Host ""
            Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
            return
        }

        # Display suggestions
        Write-Host "`n$(Get-LocalizedString 'smart_foundSuggestions' -FormatArgs @($suggestions.Count))" -ForegroundColor $Global:ColorScheme.Success
        Write-Host ""

        Write-Host (Get-LocalizedString "smart_suggestionsTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        $format = "{0,-4} {1,-12} {2,-40} {3,15} {4,10}"
        Write-Host ($format -f "#", "Action", "Rule Description", "Confidence", "Based On") -ForegroundColor $Global:ColorScheme.Header
        Write-Host ("─" * 100) -ForegroundColor $Global:ColorScheme.Border

        $index = 1
        foreach ($suggestion in $suggestions) {
            $description = if ($suggestion.Description.Length -gt 39) {
                $suggestion.Description.Substring(0, 36) + "..."
            } else {
                $suggestion.Description
            }

            $confidenceColor = if ($suggestion.Confidence -ge 80) {
                $Global:ColorScheme.Success
            } elseif ($suggestion.Confidence -ge 60) {
                $Global:ColorScheme.Info
            } else {
                $Global:ColorScheme.Warning
            }

            Write-Host ($format -f $index, $suggestion.Type, $description, "$($suggestion.Confidence)%", "$($suggestion.BasedOnActions) actions") `
                -ForegroundColor $confidenceColor
            $index++
        }

        Write-Host ""

        # Menu
        while ($true) {
            Write-Host (Get-LocalizedString "smart_menuTitle") -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  1. $(Get-LocalizedString 'smart_applyRule')" -ForegroundColor Green
            Write-Host "  2. $(Get-LocalizedString 'smart_viewActions')" -ForegroundColor Cyan
            Write-Host "  3. $(Get-LocalizedString 'smart_clearHistory')" -ForegroundColor Yellow
            Write-Host "  4. $(Get-LocalizedString 'smart_exportReport')" -ForegroundColor Magenta
            Write-Host "  Q. $(Get-LocalizedString 'unsubscribe_back')" -ForegroundColor Red
            Write-Host ""

            $choice = Read-Host (Get-LocalizedString "unsubscribe_selectAction")

            switch ($choice.ToUpper()) {
                "1" {
                    # Apply a rule
                    $ruleNum = Read-Host (Get-LocalizedString "smart_enterRuleNumber")
                    if ($ruleNum -match '^\d+$' -and [int]$ruleNum -ge 1 -and [int]$ruleNum -le $suggestions.Count) {
                        $selectedRule = $suggestions[[int]$ruleNum - 1]
                        Invoke-ApplySmartRule -UserEmail $UserEmail -Rule $selectedRule
                    } else {
                        Write-Host (Get-LocalizedString "unsubscribe_invalidNumber") -ForegroundColor $Global:ColorScheme.Warning
                    }
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "2" {
                    # View action history
                    Show-ActionHistory
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "3" {
                    # Clear history
                    $confirm = Show-Confirmation -Message (Get-LocalizedString "smart_confirmClear")
                    if ($confirm) {
                        Clear-ActionHistory
                        Write-Host (Get-LocalizedString "smart_historyCleared") -ForegroundColor $Global:ColorScheme.Success
                    }
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
                }
                "4" {
                    # Export report
                    Export-SmartOrganizerReport -Suggestions $suggestions
                    Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
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
        Write-Error "Error in smart organizer: $($_.Exception.Message)"
        Write-Host "`n$(Get-LocalizedString 'script_errorOccurred' -FormatArgs @($_.Exception.Message))" -ForegroundColor $Global:ColorScheme.Error
        Read-Host (Get-LocalizedString "mainMenu_actionPressEnterToContinue")
    }
}

# Function: Invoke-ApplySmartRule
function Invoke-ApplySmartRule {
    <#
    .SYNOPSIS
        Applies a smart rule to current emails
    .PARAMETER UserEmail
        User email address
    .PARAMETER Rule
        Rule to apply
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Rule
    )

    Write-Host ""
    Write-Host (Get-LocalizedString "smart_applyingRule") -ForegroundColor $Global:ColorScheme.Info
    Write-Host "  $($Rule.Description)" -ForegroundColor $Global:ColorScheme.Value
    Write-Host ""

    $confirm = Show-Confirmation -Message (Get-LocalizedString "smart_confirmApply")

    if (-not $confirm) {
        Write-Host (Get-LocalizedString "performActionAll_moveCancelled") -ForegroundColor $Global:ColorScheme.Warning
        return
    }

    # Get emails from this domain
    $cache = Get-SenderCache
    if (-not $cache.ContainsKey($Rule.SenderDomain)) {
        Write-Host (Get-LocalizedString "smart_noDomainEmails" -FormatArgs @($Rule.SenderDomain)) -ForegroundColor $Global:ColorScheme.Warning
        return
    }

    $emails = $cache[$Rule.SenderDomain].Messages

    Write-Host ""
    Write-Host (Get-LocalizedString "smart_applyingToEmails" -FormatArgs @($emails.Count, $Rule.SenderDomain)) -ForegroundColor $Global:ColorScheme.Info

    $successCount = 0
    $errorCount = 0

    switch ($Rule.Type) {
        "Move" {
            # Get or create folder
            $folders = Get-GraphMailFolders -UserId $UserEmail
            $targetFolder = $folders | Where-Object { $_.displayName -eq $Rule.DestinationFolder }

            if (-not $targetFolder) {
                Write-Host (Get-LocalizedString "smart_creatingFolder" -FormatArgs @($Rule.DestinationFolder)) -ForegroundColor $Global:ColorScheme.Info
                $targetFolder = New-GraphMailFolder -UserId $UserEmail -DisplayName $Rule.DestinationFolder
            }

            foreach ($email in $emails) {
                try {
                    Move-GraphMessage -UserId $UserEmail -MessageId $email.MessageId -DestinationFolderId $targetFolder.id | Out-Null
                    $successCount++
                }
                catch {
                    $errorCount++
                }
            }
        }
        "Delete" {
            foreach ($email in $emails) {
                try {
                    Remove-GraphMessage -UserId $UserEmail -MessageId $email.MessageId | Out-Null
                    $successCount++
                }
                catch {
                    $errorCount++
                }
            }
        }
        "Archive" {
            $folders = Get-GraphMailFolders -UserId $UserEmail
            $archiveFolder = $folders | Where-Object { $_.displayName -eq "Archive" }

            if (-not $archiveFolder) {
                $archiveFolder = New-GraphMailFolder -UserId $UserEmail -DisplayName "Archive"
            }

            foreach ($email in $emails) {
                try {
                    Move-GraphMessage -UserId $UserEmail -MessageId $email.MessageId -DestinationFolderId $archiveFolder.id | Out-Null
                    $successCount++
                }
                catch {
                    $errorCount++
                }
            }
        }
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "smart_ruleApplied" -FormatArgs @($successCount)) -ForegroundColor $Global:ColorScheme.Success
    if ($errorCount -gt 0) {
        Write-Host (Get-LocalizedString "performActionAll_moveErrorCount" -FormatArgs @($errorCount)) -ForegroundColor $Global:ColorScheme.Warning
    }
}

# Function: Show-ActionHistory
function Show-ActionHistory {
    <#
    .SYNOPSIS
        Shows user action history
    #>
    [CmdletBinding()]
    param()

    $data = Get-Content -Path $script:SmartRulesPath -Raw | ConvertFrom-Json

    if ($data.Actions.Count -eq 0) {
        Write-Host "`n$(Get-LocalizedString 'smart_noHistory')" -ForegroundColor $Global:ColorScheme.Warning
        return
    }

    Write-Host ""
    Write-Host (Get-LocalizedString "smart_actionHistory" -FormatArgs @($data.Actions.Count)) -ForegroundColor $Global:ColorScheme.SectionHeader
    Write-Host ("─" * 80) -ForegroundColor $Global:ColorScheme.Border

    $recent = $data.Actions | Select-Object -Last 20 | Sort-Object { ConvertTo-SafeDateTime -DateTimeValue $_.Timestamp } -Descending

    foreach ($action in $recent) {
        $timestamp = ConvertTo-SafeDateTime -DateTimeValue $action.Timestamp.ToString('yyyy-MM-dd HH:mm')
        Write-Host "  [$timestamp] " -NoNewline -ForegroundColor $Global:ColorScheme.Muted
        Write-Host "$($action.ActionType) " -NoNewline -ForegroundColor $Global:ColorScheme.Value
        Write-Host "from $($action.SenderDomain)" -NoNewline -ForegroundColor $Global:ColorScheme.Normal
        if ($action.DestinationFolder) {
            Write-Host " → $($action.DestinationFolder)" -NoNewline -ForegroundColor $Global:ColorScheme.Info
        }
        Write-Host " ($($action.EmailCount) emails)" -ForegroundColor $Global:ColorScheme.Muted
    }
}

# Function: Clear-ActionHistory
function Clear-ActionHistory {
    <#
    .SYNOPSIS
        Clears action history
    #>
    [CmdletBinding()]
    param()

    $data = Get-Content -Path $script:SmartRulesPath -Raw | ConvertFrom-Json
    $data.Actions = @()
    $data.SuggestedRules = @()
    $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:SmartRulesPath -Encoding UTF8
}

# Function: Export-SmartOrganizerReport
function Export-SmartOrganizerReport {
    <#
    .SYNOPSIS
        Exports smart organizer report
    .PARAMETER Suggestions
        Array of suggestions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Suggestions
    )

    $defaultPath = Join-Path $PSScriptRoot "..\..\smart_organizer_report.csv"
    $exportPath = Read-Host (Get-LocalizedString "unsubscribe_exportPath" -FormatArgs @($defaultPath))

    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $defaultPath
    }

    $Suggestions | Select-Object Type, SenderDomain, DestinationFolder, Confidence, BasedOnActions, Description |
        Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host (Get-LocalizedString "unsubscribe_exportSuccess" -FormatArgs @($exportPath)) -ForegroundColor $Global:ColorScheme.Success
}

# Export functions
Export-ModuleMember -Function Show-SmartOrganizer, Track-UserAction, Analyze-UserPatterns, `
    Initialize-SmartOrganizer
