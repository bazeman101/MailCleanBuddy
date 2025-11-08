<#
.SYNOPSIS
    Rule Engine for MailCleanBuddy Email Automation
.DESCRIPTION
    Provides if-then rule automation for automatic email management
    Supports conditions, actions, scheduling, and audit logging
#>

# Script-level variables
$Script:RulesPath = $null
$Script:AuditLogPath = $null

<#
.SYNOPSIS
    Initializes the rule engine
#>
function Initialize-RuleEngine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        $sanitizedEmail = $UserEmail -replace '[\\/:*?"<>|]', '_'
        $homeDir = if ($IsWindows -or $null -eq $IsWindows) { $env:USERPROFILE } else { $env:HOME }
        $ruleDir = Join-Path $homeDir ".mailcleanbuddy"

        if (-not (Test-Path $ruleDir)) {
            New-Item -Path $ruleDir -ItemType Directory -Force | Out-Null
        }

        $Script:RulesPath = Join-Path $ruleDir "automation_rules_$sanitizedEmail.json"
        $Script:AuditLogPath = Join-Path $ruleDir "rule_audit_$sanitizedEmail.log"

        if (-not (Test-Path $Script:RulesPath)) {
            $initialData = @{
                Version = "1.0"
                Rules = @()
                LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
            $initialData | ConvertTo-Json -Depth 10 | Set-Content -Path $Script:RulesPath -Encoding UTF8
        }

        return $true
    } catch {
        Write-Warning "Failed to initialize rule engine: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Creates a new automation rule
#>
function New-AutomationRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string]$Description,

        [Parameter(Mandatory = $true)]
        [hashtable]$Conditions,

        [Parameter(Mandatory = $true)]
        [hashtable]$Action,

        [Parameter(Mandatory = $false)]
        [bool]$Enabled = $true,

        [Parameter(Mandatory = $false)]
        [int]$Priority = 5
    )

    return [PSCustomObject]@{
        Id = [guid]::NewGuid().ToString()
        Name = $Name
        Description = $Description
        Conditions = $Conditions
        Action = $Action
        Enabled = $Enabled
        Priority = $Priority
        CreatedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        LastExecuted = $null
        ExecutionCount = 0
        SuccessCount = 0
        FailureCount = 0
    }
}

<#
.SYNOPSIS
    Tests if a message matches rule conditions
#>
function Test-RuleConditions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Message,

        [Parameter(Mandatory = $true)]
        [hashtable]$Conditions
    )

    $results = @()

    # From/Sender condition
    if ($Conditions.ContainsKey('From') -and $Conditions.From) {
        $senderEmail = if ($Message.PSObject.Properties['SenderEmailAddress']) {
            $Message.SenderEmailAddress
        } else {
            ""
        }
        $matches = $senderEmail -like "*$($Conditions.From)*"
        $results += $matches
    }

    # Subject condition
    if ($Conditions.ContainsKey('SubjectContains') -and $Conditions.SubjectContains) {
        $subject = if ($Message.PSObject.Properties['Subject']) { $Message.Subject } else { "" }
        $matches = $subject -like "*$($Conditions.SubjectContains)*"
        $results += $matches
    }

    # Has Attachments condition
    if ($Conditions.ContainsKey('HasAttachments')) {
        $hasAttach = if ($Message.PSObject.Properties['HasAttachments']) { $Message.HasAttachments } else { $false }
        $results += ($hasAttach -eq $Conditions.HasAttachments)
    }

    # Is Read condition
    if ($Conditions.ContainsKey('IsRead')) {
        $isRead = if ($Message.PSObject.Properties['IsRead']) { $Message.IsRead } else { $false }
        $results += ($isRead -eq $Conditions.IsRead)
    }

    # Importance condition
    if ($Conditions.ContainsKey('Importance') -and $Conditions.Importance) {
        $importance = if ($Message.PSObject.Properties['Importance']) { $Message.Importance } else { "Normal" }
        $results += ($importance -eq $Conditions.Importance)
    }

    # Size condition
    if ($Conditions.ContainsKey('MinSize') -and $Conditions.MinSize -gt 0) {
        $size = if ($Message.PSObject.Properties['Size']) { $Message.Size } else { 0 }
        $results += ($size -ge $Conditions.MinSize)
    }

    if ($Conditions.ContainsKey('MaxSize') -and $Conditions.MaxSize -gt 0) {
        $size = if ($Message.PSObject.Properties['Size']) { $Message.Size } else { 0 }
        $results += ($size -le $Conditions.MaxSize)
    }

    # Date condition (older than X days)
    if ($Conditions.ContainsKey('OlderThanDays') -and $Conditions.OlderThanDays -gt 0) {
        $receivedDate = ConvertTo-SafeDateTime -DateTimeValue $Message.ReceivedDateTime
        if ($receivedDate) {
            $daysDiff = ((Get-Date) - $receivedDate).Days
            $results += ($daysDiff -ge $Conditions.OlderThanDays)
        }
    }

    # Body contains condition
    if ($Conditions.ContainsKey('BodyContains') -and $Conditions.BodyContains) {
        $body = if ($Message.PSObject.Properties['BodyPreview']) { $Message.BodyPreview } else { "" }
        $matches = $body -like "*$($Conditions.BodyContains)*"
        $results += $matches
    }

    # Category condition
    if ($Conditions.ContainsKey('Category') -and $Conditions.Category) {
        $categories = if ($Message.PSObject.Properties['Categories']) { $Message.Categories } else { @() }
        $matches = $categories -contains $Conditions.Category
        $results += $matches
    }

    # Logic operator (default AND)
    $operator = if ($Conditions.ContainsKey('Operator')) { $Conditions.Operator } else { "And" }

    if ($results.Count -eq 0) {
        return $false
    }

    if ($operator -eq "Or") {
        return ($results -contains $true)
    } else {
        return ($results -notcontains $false)
    }
}

<#
.SYNOPSIS
    Executes a rule action on a message
#>
function Invoke-RuleAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        $Message,

        [Parameter(Mandatory = $true)]
        [hashtable]$Action,

        [Parameter(Mandatory = $false)]
        [switch]$DryRun
    )

    try {
        $actionType = $Action.Type
        $msgId = if ($Message.Id) { $Message.Id } elseif ($Message.MessageId) { $Message.MessageId } else { $null }

        if (-not $msgId) {
            return @{ Success = $false; Error = "Message ID not found" }
        }

        if ($DryRun) {
            $actionDesc = switch ($actionType) {
                "Delete" { "DELETE message" }
                "Move" { "MOVE to folder: $($Action.FolderId)" }
                "MarkAsRead" { "MARK AS READ" }
                "MarkAsUnread" { "MARK AS UNREAD" }
                "Categorize" { "ADD CATEGORY: $($Action.Category)" }
                "Flag" { "FLAG message" }
                default { "UNKNOWN ACTION: $actionType" }
            }

            Write-LogMessage -Level "Info" -Message "DRY RUN: Would execute $actionDesc on message '$($Message.Subject)'" -Source "RuleEngine"
            return @{ Success = $true; DryRun = $true; Action = $actionDesc }
        }

        switch ($actionType) {
            "Delete" {
                Remove-MgUserMessage -UserId $UserEmail -MessageId $msgId -ErrorAction Stop
                Write-LogMessage -Level "Info" -Message "Rule action: Deleted message '$($Message.Subject)'" -Source "RuleEngine"
                return @{ Success = $true; Action = "Deleted" }
            }

            "Move" {
                if (-not $Action.FolderId) {
                    return @{ Success = $false; Error = "FolderId not specified" }
                }
                Move-MgUserMessage -UserId $UserEmail -MessageId $msgId -DestinationId $Action.FolderId -ErrorAction Stop
                Write-LogMessage -Level "Info" -Message "Rule action: Moved message '$($Message.Subject)' to folder $($Action.FolderId)" -Source "RuleEngine"
                return @{ Success = $true; Action = "Moved" }
            }

            "MarkAsRead" {
                $body = @{ IsRead = $true } | ConvertTo-Json
                Update-MgUserMessage -UserId $UserEmail -MessageId $msgId -BodyParameter $body -ErrorAction Stop
                Write-LogMessage -Level "Info" -Message "Rule action: Marked message as read '$($Message.Subject)'" -Source "RuleEngine"
                return @{ Success = $true; Action = "MarkedAsRead" }
            }

            "MarkAsUnread" {
                $body = @{ IsRead = $false } | ConvertTo-Json
                Update-MgUserMessage -UserId $UserEmail -MessageId $msgId -BodyParameter $body -ErrorAction Stop
                Write-LogMessage -Level "Info" -Message "Rule action: Marked message as unread '$($Message.Subject)'" -Source "RuleEngine"
                return @{ Success = $true; Action = "MarkedAsUnread" }
            }

            "Categorize" {
                if (-not $Action.Category) {
                    return @{ Success = $false; Error = "Category not specified" }
                }
                $existingCategories = if ($Message.PSObject.Properties['Categories']) { $Message.Categories } else { @() }
                $newCategories = $existingCategories + $Action.Category
                $body = @{ Categories = $newCategories } | ConvertTo-Json
                Update-MgUserMessage -UserId $UserEmail -MessageId $msgId -BodyParameter $body -ErrorAction Stop
                Write-LogMessage -Level "Info" -Message "Rule action: Added category '$($Action.Category)' to message '$($Message.Subject)'" -Source "RuleEngine"
                return @{ Success = $true; Action = "Categorized" }
            }

            "Flag" {
                $body = @{
                    Flag = @{
                        FlagStatus = "Flagged"
                    }
                } | ConvertTo-Json -Depth 10
                Update-MgUserMessage -UserId $UserEmail -MessageId $msgId -BodyParameter $body -ErrorAction Stop
                Write-LogMessage -Level "Info" -Message "Rule action: Flagged message '$($Message.Subject)'" -Source "RuleEngine"
                return @{ Success = $true; Action = "Flagged" }
            }

            default {
                return @{ Success = $false; Error = "Unknown action type: $actionType" }
            }
        }
    } catch {
        Write-LogMessage -Level "Error" -Message "Failed to execute rule action: $($_.Exception.Message)" -Exception $_.Exception -Source "RuleEngine"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

<#
.SYNOPSIS
    Executes all enabled rules on messages
#>
function Invoke-RuleExecution {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Messages,

        [Parameter(Mandatory = $false)]
        [switch]$DryRun
    )

    try {
        # Check if rules engine is enabled
        $rulesEnabled = Get-ConfigValue -Path "Rules.Enabled" -DefaultValue $true
        if (-not $rulesEnabled) {
            Write-Host "Rule engine is disabled in config." -ForegroundColor $Global:ColorScheme.Warning
            return @{ TotalRules = 0; ProcessedMessages = 0; ActionsExecuted = 0; Errors = 0 }
        }

        Initialize-RuleEngine -UserEmail $UserEmail | Out-Null

        # Load rules
        $data = Get-Content -Path $Script:RulesPath -Raw | ConvertFrom-Json
        $rules = $data.Rules | Where-Object { $_.Enabled -eq $true } | Sort-Object Priority -Descending

        if ($rules.Count -eq 0) {
            Write-Host "No enabled rules found." -ForegroundColor $Global:ColorScheme.Info
            return @{ TotalRules = 0; ProcessedMessages = 0; ActionsExecuted = 0; Errors = 0 }
        }

        Write-Host "`n⚙️  Executing Rules..." -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host "Found $($rules.Count) enabled rule(s)" -ForegroundColor $Global:ColorScheme.Info
        if ($DryRun) {
            Write-Host "🔍 DRY RUN MODE - No changes will be made" -ForegroundColor $Global:ColorScheme.Warning
        }
        Write-Host ""

        $stats = @{
            TotalRules = $rules.Count
            ProcessedMessages = 0
            ActionsExecuted = 0
            Errors = 0
            RuleStats = @{}
        }

        foreach ($rule in $rules) {
            $stats.RuleStats[$rule.Name] = @{
                Matched = 0
                Executed = 0
                Failed = 0
            }
        }

        $processed = 0
        foreach ($message in $Messages) {
            $processed++

            if ($processed % 20 -eq 0) {
                $percent = [Math]::Round(($processed / $Messages.Count) * 100)
                Write-Progress -Activity "Processing rules..." -Status "Message $processed of $($Messages.Count)" -PercentComplete $percent
            }

            foreach ($rule in $rules) {
                # Test if message matches rule conditions
                $matches = Test-RuleConditions -Message $message -Conditions $rule.Conditions

                if ($matches) {
                    $stats.RuleStats[$rule.Name].Matched++

                    # Execute action
                    $result = Invoke-RuleAction -UserEmail $UserEmail -Message $message -Action $rule.Action -DryRun:$DryRun

                    if ($result.Success) {
                        $stats.ActionsExecuted++
                        $stats.RuleStats[$rule.Name].Executed++

                        # Log to audit
                        Write-RuleAuditLog -RuleName $rule.Name -Message $message -Action $rule.Action.Type -Result "Success" -DryRun:$DryRun

                        # Update rule statistics
                        $rule.ExecutionCount++
                        $rule.SuccessCount++
                        $rule.LastExecuted = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    } else {
                        $stats.Errors++
                        $stats.RuleStats[$rule.Name].Failed++

                        Write-RuleAuditLog -RuleName $rule.Name -Message $message -Action $rule.Action.Type -Result "Failed: $($result.Error)" -DryRun:$DryRun
                        Write-LogMessage -Level "Error" -Message "Rule '$($rule.Name)' failed: $($result.Error)" -Source "RuleEngine"

                        $rule.ExecutionCount++
                        $rule.FailureCount++
                    }

                    # One rule per message (unless specified otherwise)
                    break
                }
            }

            $stats.ProcessedMessages++
        }

        Write-Progress -Activity "Processing rules..." -Completed

        # Save updated rule statistics
        if (-not $DryRun) {
            $data.Rules = $rules
            $data.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $data | ConvertTo-Json -Depth 10 | Set-Content -Path $Script:RulesPath -Encoding UTF8
        }

        return $stats
    } catch {
        Write-Error "Error executing rules: $($_.Exception.Message)"
        Write-LogMessage -Level "Error" -Message "Rule execution failed: $($_.Exception.Message)" -Exception $_.Exception -Source "RuleEngine"
        return @{ TotalRules = 0; ProcessedMessages = 0; ActionsExecuted = 0; Errors = 1 }
    }
}

<#
.SYNOPSIS
    Writes to rule audit log
#>
function Write-RuleAuditLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RuleName,

        [Parameter(Mandatory = $true)]
        $Message,

        [Parameter(Mandatory = $true)]
        [string]$Action,

        [Parameter(Mandatory = $true)]
        [string]$Result,

        [Parameter(Mandatory = $false)]
        [switch]$DryRun
    )

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $dryRunPrefix = if ($DryRun) { "[DRY RUN] " } else { "" }
        $subject = if ($Message.Subject) { $Message.Subject } else { "(No Subject)" }

        $logEntry = "$timestamp | ${dryRunPrefix}Rule: $RuleName | Action: $Action | Message: $subject | Result: $Result"

        Add-Content -Path $Script:AuditLogPath -Value $logEntry -Encoding UTF8
    } catch {
        Write-Verbose "Failed to write audit log: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Saves an automation rule
#>
function Save-AutomationRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Rule
    )

    try {
        if (-not $Script:RulesPath) {
            throw "Rule engine not initialized"
        }

        $data = Get-Content -Path $Script:RulesPath -Raw | ConvertFrom-Json -AsHashtable

        # Check for duplicate ID
        $existingIndex = -1
        for ($i = 0; $i -lt $data.Rules.Count; $i++) {
            if ($data.Rules[$i].Id -eq $Rule.Id) {
                $existingIndex = $i
                break
            }
        }

        if ($existingIndex -ge 0) {
            # Update existing rule
            $data.Rules[$existingIndex] = $Rule
            Write-Verbose "Updated existing rule: $($Rule.Name)"
        } else {
            # Add new rule
            $data.Rules += $Rule
            Write-Verbose "Added new rule: $($Rule.Name)"
        }

        $data.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $Script:RulesPath -Encoding UTF8

        Write-LogMessage -Level "Info" -Message "Saved automation rule: $($Rule.Name)" -Source "RuleEngine"
        return $true
    } catch {
        Write-Error "Failed to save rule: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Gets all automation rules
#>
function Get-AutomationRules {
    [CmdletBinding()]
    param()

    try {
        if (-not $Script:RulesPath -or -not (Test-Path $Script:RulesPath)) {
            return @()
        }

        $data = Get-Content -Path $Script:RulesPath -Raw | ConvertFrom-Json
        return $data.Rules
    } catch {
        Write-Warning "Failed to load automation rules: $($_.Exception.Message)"
        return @()
    }
}

<#
.SYNOPSIS
    Removes an automation rule
#>
function Remove-AutomationRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RuleId
    )

    try {
        if (-not $Script:RulesPath) {
            throw "Rule engine not initialized"
        }

        $data = Get-Content -Path $Script:RulesPath -Raw | ConvertFrom-Json -AsHashtable
        $data.Rules = $data.Rules | Where-Object { $_.Id -ne $RuleId }
        $data.LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $Script:RulesPath -Encoding UTF8

        Write-LogMessage -Level "Info" -Message "Removed automation rule: $RuleId" -Source "RuleEngine"
        return $true
    } catch {
        Write-Error "Failed to remove rule: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Shows rule management UI
#>
function Show-RuleManagement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Initialize-RuleEngine -UserEmail $UserEmail | Out-Null

        while ($true) {
            Clear-Host
            Write-Host "`n⚙️  Rule Management" -ForegroundColor $Global:ColorScheme.Highlight
            Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
            Write-Host ""

            $rules = Get-AutomationRules

            Write-Host "📋 Automation Rules ($($rules.Count))" -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host ""

            if ($rules.Count -eq 0) {
                Write-Host "  No rules configured yet." -ForegroundColor $Global:ColorScheme.Muted
            } else {
                $index = 1
                foreach ($rule in $rules | Sort-Object Priority -Descending) {
                    $statusIcon = if ($rule.Enabled) { "✓" } else { "✗" }
                    $statusColor = if ($rule.Enabled) { $Global:ColorScheme.Success } else { $Global:ColorScheme.Muted }

                    Write-Host "  $index. " -NoNewline -ForegroundColor $Global:ColorScheme.Muted
                    Write-Host "[$statusIcon] " -NoNewline -ForegroundColor $statusColor
                    Write-Host "$($rule.Name) " -NoNewline -ForegroundColor $Global:ColorScheme.Value
                    Write-Host "(Priority: $($rule.Priority))" -ForegroundColor $Global:ColorScheme.Muted

                    if ($rule.Description) {
                        Write-Host "     $($rule.Description)" -ForegroundColor $Global:ColorScheme.Info
                    }

                    Write-Host "     Action: " -NoNewline -ForegroundColor $Global:ColorScheme.Muted
                    Write-Host "$($rule.Action.Type)" -NoNewline -ForegroundColor $Global:ColorScheme.Highlight
                    Write-Host " | Executed: $($rule.ExecutionCount) times (✓$($rule.SuccessCount) ✗$($rule.FailureCount))" -ForegroundColor $Global:ColorScheme.Muted
                    Write-Host ""

                    $index++
                }
            }

            Write-Host "Actions:" -ForegroundColor $Global:ColorScheme.SectionHeader
            Write-Host "  [N] Create new rule" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [E] Enable/Disable rule" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [D] Delete rule" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [T] Test rules (dry run)" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [A] View audit log" -ForegroundColor $Global:ColorScheme.Info
            Write-Host "  [Q] Back" -ForegroundColor $Global:ColorScheme.Info
            Write-Host ""

            $choice = Read-Host "Select action"

            switch ($choice.ToUpper()) {
                "N" {
                    $newRule = Show-RuleBuilder -UserEmail $UserEmail
                    if ($newRule) {
                        Save-AutomationRule -Rule $newRule
                        Write-Host "Rule created successfully!" -ForegroundColor $Global:ColorScheme.Success
                        Start-Sleep -Seconds 2
                    }
                }
                "E" {
                    if ($rules.Count -gt 0) {
                        $ruleNum = Read-Host "Enter rule number to toggle"
                        if ($ruleNum -match '^\d+$' -and [int]$ruleNum -ge 1 -and [int]$ruleNum -le $rules.Count) {
                            $selectedRule = $rules[[int]$ruleNum - 1]
                            $selectedRule.Enabled = -not $selectedRule.Enabled
                            Save-AutomationRule -Rule $selectedRule
                            $status = if ($selectedRule.Enabled) { "enabled" } else { "disabled" }
                            Write-Host "Rule $status!" -ForegroundColor $Global:ColorScheme.Success
                            Start-Sleep -Seconds 1
                        }
                    }
                }
                "D" {
                    if ($rules.Count -gt 0) {
                        $ruleNum = Read-Host "Enter rule number to delete"
                        if ($ruleNum -match '^\d+$' -and [int]$ruleNum -ge 1 -and [int]$ruleNum -le $rules.Count) {
                            $selectedRule = $rules[[int]$ruleNum - 1]
                            $confirm = Read-Host "Delete rule '$($selectedRule.Name)'? (yes/no)"
                            if ($confirm -eq "yes" -or $confirm -eq "y") {
                                Remove-AutomationRule -RuleId $selectedRule.Id
                                Write-Host "Rule deleted!" -ForegroundColor $Global:ColorScheme.Success
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                }
                "T" {
                    Write-Host "`nTesting rules on cached messages..." -ForegroundColor $Global:ColorScheme.Info
                    $cache = Get-SenderCache
                    $testMessages = @()
                    $count = 0
                    foreach ($domain in $cache.Keys) {
                        foreach ($msg in $cache[$domain].Messages) {
                            $testMessages += $msg
                            $count++
                            if ($count -ge 50) { break }  # Limit to 50 messages for testing
                        }
                        if ($count -ge 50) { break }
                    }

                    if ($testMessages.Count -gt 0) {
                        $result = Invoke-RuleExecution -UserEmail $UserEmail -Messages $testMessages -DryRun
                        Write-Host "`nDry Run Results:" -ForegroundColor $Global:ColorScheme.Success
                        Write-Host "  Tested $($result.ProcessedMessages) messages" -ForegroundColor $Global:ColorScheme.Info
                        Write-Host "  Actions would execute: $($result.ActionsExecuted)" -ForegroundColor $Global:ColorScheme.Info
                        Write-Host "  Errors: $($result.Errors)" -ForegroundColor $Global:ColorScheme.Info
                    } else {
                        Write-Host "No messages in cache for testing." -ForegroundColor $Global:ColorScheme.Warning
                    }
                    Read-Host "`nPress Enter to continue"
                }
                "A" {
                    Show-RuleAuditLog
                    Read-Host "`nPress Enter to continue"
                }
                "Q" {
                    return
                }
            }
        }
    } catch {
        Write-Error "Error in rule management: $($_.Exception.Message)"
        Read-Host "Press Enter to continue"
    }
}

<#
.SYNOPSIS
    Interactive rule builder
#>
function Show-RuleBuilder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    try {
        Clear-Host
        Write-Host "`n⚙️  Rule Builder" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        # Rule name
        $ruleName = Read-Host "Rule name (required)"
        if ([string]::IsNullOrWhiteSpace($ruleName)) {
            Write-Host "Rule name is required." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return $null
        }

        $description = Read-Host "Description (optional)"

        # Build conditions
        Write-Host "`n📋 Conditions (IF...)" -ForegroundColor $Global:ColorScheme.SectionHeader
        $conditions = @{}

        $from = Read-Host "From (sender email/domain, or skip)"
        if (-not [string]::IsNullOrWhiteSpace($from)) { $conditions.From = $from }

        $subjectContains = Read-Host "Subject contains (or skip)"
        if (-not [string]::IsNullOrWhiteSpace($subjectContains)) { $conditions.SubjectContains = $subjectContains }

        $hasAttach = Read-Host "Has attachments? (yes/no/skip)"
        if ($hasAttach -eq "yes" -or $hasAttach -eq "y") { $conditions.HasAttachments = $true }
        elseif ($hasAttach -eq "no" -or $hasAttach -eq "n") { $conditions.HasAttachments = $false }

        $olderThan = Read-Host "Older than X days (or skip)"
        if ($olderThan -match '^\d+$') { $conditions.OlderThanDays = [int]$olderThan }

        if ($conditions.Count -eq 0) {
            Write-Host "At least one condition is required." -ForegroundColor $Global:ColorScheme.Warning
            Read-Host "Press Enter to continue"
            return $null
        }

        # Build action
        Write-Host "`n⚡ Action (THEN...)" -ForegroundColor $Global:ColorScheme.SectionHeader
        Write-Host "  1. Delete message" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  2. Move to folder" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  3. Mark as read" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  4. Mark as unread" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  5. Add category" -ForegroundColor $Global:ColorScheme.Info
        Write-Host "  6. Flag message" -ForegroundColor $Global:ColorScheme.Info

        $actionChoice = Read-Host "Select action (1-6)"

        $action = @{}
        switch ($actionChoice) {
            "1" { $action.Type = "Delete" }
            "2" {
                $action.Type = "Move"
                $folderId = Read-Host "Enter folder ID"
                $action.FolderId = $folderId
            }
            "3" { $action.Type = "MarkAsRead" }
            "4" { $action.Type = "MarkAsUnread" }
            "5" {
                $action.Type = "Categorize"
                $category = Read-Host "Enter category name"
                $action.Category = $category
            }
            "6" { $action.Type = "Flag" }
            default {
                Write-Host "Invalid action choice." -ForegroundColor $Global:ColorScheme.Warning
                Read-Host "Press Enter to continue"
                return $null
            }
        }

        # Priority
        $priorityInput = Read-Host "Priority (1-10, default 5)"
        $priority = if ($priorityInput -match '^\d+$') { [int]$priorityInput } else { 5 }

        # Create rule
        $rule = New-AutomationRule -Name $ruleName `
                                   -Description $description `
                                   -Conditions $conditions `
                                   -Action $action `
                                   -Priority $priority

        Write-Host "`n✅ Rule created successfully!" -ForegroundColor $Global:ColorScheme.Success
        return $rule
    } catch {
        Write-Error "Error in rule builder: $($_.Exception.Message)"
        Read-Host "Press Enter to continue"
        return $null
    }
}

<#
.SYNOPSIS
    Shows the rule audit log
#>
function Show-RuleAuditLog {
    [CmdletBinding()]
    param()

    try {
        Clear-Host
        Write-Host "`n📜 Rule Audit Log" -ForegroundColor $Global:ColorScheme.Highlight
        Write-Host ("=" * 100) -ForegroundColor $Global:ColorScheme.Border
        Write-Host ""

        if (-not (Test-Path $Script:AuditLogPath)) {
            Write-Host "No audit log found." -ForegroundColor $Global:ColorScheme.Muted
            return
        }

        $logEntries = Get-Content -Path $Script:AuditLogPath -Tail 50

        foreach ($entry in $logEntries) {
            Write-Host $entry -ForegroundColor $Global:ColorScheme.Normal
        }

        Write-Host ""
        Write-Host "Showing last 50 entries" -ForegroundColor $Global:ColorScheme.Muted
    } catch {
        Write-Error "Error showing audit log: $($_.Exception.Message)"
    }
}

# Export functions
Export-ModuleMember -Function Initialize-RuleEngine, New-AutomationRule, Test-RuleConditions, `
                              Invoke-RuleAction, Invoke-RuleExecution, Save-AutomationRule, `
                              Get-AutomationRules, Remove-AutomationRule, Show-RuleManagement, `
                              Show-RuleBuilder, Show-RuleAuditLog
