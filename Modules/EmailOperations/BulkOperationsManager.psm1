<#
.SYNOPSIS
    Bulk Operations Manager for MailCleanBuddy
.DESCRIPTION
    Provides optimized bulk operations with parallel processing and progress tracking
#>

<#
.SYNOPSIS
    Processes bulk delete with parallel processing
.PARAMETER UserEmail
    User email address
.PARAMETER Messages
    Array of messages to delete
.PARAMETER ShowProgress
    Show progress bar
#>
function Invoke-BulkDelete {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Messages,

        [Parameter(Mandatory = $false)]
        [switch]$ShowProgress
    )

    if ($Messages.Count -eq 0) {
        return @{ Success = 0; Failed = 0; Errors = @() }
    }

    # Get config settings
    $enableParallel = Get-ConfigValue -Path "BulkOperations.EnableParallelProcessing" -DefaultValue $true
    $maxThreads = Get-ConfigValue -Path "BulkOperations.MaxParallelThreads" -DefaultValue 4
    $batchSize = Get-ConfigValue -Path "BulkOperations.BatchSize" -DefaultValue 50

    $successCount = 0
    $failedCount = 0
    $errors = @()

    try {
        if ($ShowProgress) {
            Write-Progress -Activity "Deleting Emails" -Status "Preparing..." -PercentComplete 0
        }

        # Check PowerShell version for parallel support
        $canUseParallel = $PSVersionTable.PSVersion.Major -ge 7 -and $enableParallel

        if ($canUseParallel) {
            # Parallel processing (PS 7+)
            Write-Verbose "Using parallel processing with $maxThreads threads"

            $results = $Messages | ForEach-Object -Parallel {
                try {
                    Remove-MgUserMessage -UserId $using:UserEmail -MessageId $_.Id -ErrorAction Stop
                    return @{ Success = $true; MessageId = $_.Id }
                } catch {
                    return @{ Success = $false; MessageId = $_.Id; Error = $_.Exception.Message }
                }
            } -ThrottleLimit $maxThreads

            # Process results
            $processed = 0
            foreach ($result in $results) {
                $processed++
                if ($result.Success) {
                    $successCount++
                } else {
                    $failedCount++
                    $errors += $result.Error
                }

                if ($ShowProgress) {
                    $percentComplete = [Math]::Min(100, [Math]::Round(($processed / $Messages.Count) * 100))
                    Write-Progress -Activity "Deleting Emails" -Status "Deleted $processed of $($Messages.Count)" -PercentComplete $percentComplete
                }
            }
        } else {
            # Sequential processing with batches
            Write-Verbose "Using sequential batch processing"

            $processed = 0
            for ($i = 0; $i -lt $Messages.Count; $i += $batchSize) {
                $batch = $Messages[$i..[Math]::Min($i + $batchSize - 1, $Messages.Count - 1)]

                foreach ($message in $batch) {
                    try {
                        Remove-MgUserMessage -UserId $UserEmail -MessageId $message.Id -ErrorAction Stop
                        $successCount++
                    } catch {
                        $failedCount++
                        $errors += $_.Exception.Message
                    }

                    $processed++
                    if ($ShowProgress) {
                        $percentComplete = [Math]::Min(100, [Math]::Round(($processed / $Messages.Count) * 100))
                        Write-Progress -Activity "Deleting Emails" -Status "Deleted $processed of $($Messages.Count)" -PercentComplete $percentComplete
                    }
                }

                # Small delay between batches to avoid API throttling
                if ($i + $batchSize < $Messages.Count) {
                    Start-Sleep -Milliseconds (Get-ConfigValue -Path "Performance.ApiThrottleDelay" -DefaultValue 100)
                }
            }
        }

        if ($ShowProgress) {
            Write-Progress -Activity "Deleting Emails" -Completed
        }

        return @{
            Success = $successCount
            Failed = $failedCount
            Errors = $errors
        }
    } catch {
        Write-LogMessage -Level "Error" -Message "Bulk delete failed" -Exception $_.Exception -Source "BulkOperationsManager"
        if ($ShowProgress) {
            Write-Progress -Activity "Deleting Emails" -Completed
        }
        return @{
            Success = $successCount
            Failed = $failedCount + ($Messages.Count - $successCount)
            Errors = $errors + @($_.Exception.Message)
        }
    }
}

<#
.SYNOPSIS
    Processes bulk move with parallel processing
.PARAMETER UserEmail
    User email address
.PARAMETER Messages
    Array of messages to move
.PARAMETER DestinationFolderId
    Destination folder ID
.PARAMETER ShowProgress
    Show progress bar
#>
function Invoke-BulkMove {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserEmail,

        [Parameter(Mandatory = $true)]
        [array]$Messages,

        [Parameter(Mandatory = $true)]
        [string]$DestinationFolderId,

        [Parameter(Mandatory = $false)]
        [switch]$ShowProgress
    )

    if ($Messages.Count -eq 0) {
        return @{ Success = 0; Failed = 0; Errors = @() }
    }

    # Get config settings
    $enableParallel = Get-ConfigValue -Path "BulkOperations.EnableParallelProcessing" -DefaultValue $true
    $maxThreads = Get-ConfigValue -Path "BulkOperations.MaxParallelThreads" -DefaultValue 4
    $batchSize = Get-ConfigValue -Path "BulkOperations.BatchSize" -DefaultValue 50

    $successCount = 0
    $failedCount = 0
    $errors = @()

    try {
        if ($ShowProgress) {
            Write-Progress -Activity "Moving Emails" -Status "Preparing..." -PercentComplete 0
        }

        # Check PowerShell version for parallel support
        $canUseParallel = $PSVersionTable.PSVersion.Major -ge 7 -and $enableParallel

        if ($canUseParallel) {
            # Parallel processing (PS 7+)
            Write-Verbose "Using parallel processing with $maxThreads threads"

            $results = $Messages | ForEach-Object -Parallel {
                try {
                    Move-MgUserMessage -UserId $using:UserEmail -MessageId $_.Id -DestinationId $using:DestinationFolderId -ErrorAction Stop
                    return @{ Success = $true; MessageId = $_.Id }
                } catch {
                    return @{ Success = $false; MessageId = $_.Id; Error = $_.Exception.Message }
                }
            } -ThrottleLimit $maxThreads

            # Process results
            $processed = 0
            foreach ($result in $results) {
                $processed++
                if ($result.Success) {
                    $successCount++
                } else {
                    $failedCount++
                    $errors += $result.Error
                }

                if ($ShowProgress) {
                    $percentComplete = [Math]::Min(100, [Math]::Round(($processed / $Messages.Count) * 100))
                    Write-Progress -Activity "Moving Emails" -Status "Moved $processed of $($Messages.Count)" -PercentComplete $percentComplete
                }
            }
        } else {
            # Sequential processing with batches
            Write-Verbose "Using sequential batch processing"

            $processed = 0
            for ($i = 0; $i -lt $Messages.Count; $i += $batchSize) {
                $batch = $Messages[$i..[Math]::Min($i + $batchSize - 1, $Messages.Count - 1)]

                foreach ($message in $batch) {
                    try {
                        Move-MgUserMessage -UserId $UserEmail -MessageId $message.Id -DestinationId $DestinationFolderId -ErrorAction Stop
                        $successCount++
                    } catch {
                        $failedCount++
                        $errors += $_.Exception.Message
                    }

                    $processed++
                    if ($ShowProgress) {
                        $percentComplete = [Math]::Min(100, [Math]::Round(($processed / $Messages.Count) * 100))
                        Write-Progress -Activity "Moving Emails" -Status "Moved $processed of $($Messages.Count)" -PercentComplete $percentComplete
                    }
                }

                # Small delay between batches to avoid API throttling
                if ($i + $batchSize < $Messages.Count) {
                    Start-Sleep -Milliseconds (Get-ConfigValue -Path "Performance.ApiThrottleDelay" -DefaultValue 100)
                }
            }
        }

        if ($ShowProgress) {
            Write-Progress -Activity "Moving Emails" -Completed
        }

        return @{
            Success = $successCount
            Failed = $failedCount
            Errors = $errors
        }
    } catch {
        Write-LogMessage -Level "Error" -Message "Bulk move failed" -Exception $_.Exception -Source "BulkOperationsManager"
        if ($ShowProgress) {
            Write-Progress -Activity "Moving Emails" -Completed
        }
        return @{
            Success = $successCount
            Failed = $failedCount + ($Messages.Count - $successCount)
            Errors = $errors + @($_.Exception.Message)
        }
    }
}

<#
.SYNOPSIS
    Processes bulk operation with retry logic
.PARAMETER Operation
    Script block to execute for each item
.PARAMETER Items
    Items to process
.PARAMETER MaxRetries
    Maximum number of retries per item
.PARAMETER RetryDelay
    Delay between retries in milliseconds
#>
function Invoke-BulkOperationWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation,

        [Parameter(Mandatory = $true)]
        [array]$Items,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [int]$RetryDelay = 1000
    )

    $results = @()

    foreach ($item in $Items) {
        $attempts = 0
        $success = $false

        while ($attempts -lt $MaxRetries -and -not $success) {
            try {
                $result = & $Operation $item
                $results += @{ Item = $item; Success = $true; Result = $result }
                $success = $true
            } catch {
                $attempts++
                if ($attempts -ge $MaxRetries) {
                    $results += @{ Item = $item; Success = $false; Error = $_.Exception.Message; Attempts = $attempts }
                    Write-Verbose "Failed after $attempts attempts: $($_.Exception.Message)"
                } else {
                    Write-Verbose "Retry attempt $attempts for item: $item"
                    Start-Sleep -Milliseconds $RetryDelay
                }
            }
        }
    }

    return $results
}

# Export functions
Export-ModuleMember -Function Invoke-BulkDelete, Invoke-BulkMove, Invoke-BulkOperationWithRetry
