# MailCleanBuddy Module Diagnostic Script
# Run this to identify module loading issues

Write-Host "=== MailCleanBuddy Module Diagnostic ===" -ForegroundColor Cyan
Write-Host ""

$ModulesPath = Join-Path $PSScriptRoot "Modules"
$errors = @()
$warnings = @()
$success = @()

# Test Core Modules
Write-Host "Testing Core Modules..." -ForegroundColor Yellow
$coreModules = @(
    "Utilities\Localization.psm1",
    "Utilities\Helpers.psm1",
    "UI\ColorScheme.psm1",
    "UI\Display.psm1",
    "UI\MenuSystem.psm1",
    "Core\GraphApiService.psm1",
    "Core\CacheManager.psm1"
)

foreach ($module in $coreModules) {
    $modulePath = Join-Path $ModulesPath $module
    try {
        Write-Host "  Loading $module..." -NoNewline
        Import-Module $modulePath -Force -ErrorAction Stop
        Write-Host " OK" -ForegroundColor Green
        $success += $module
    } catch {
        Write-Host " FAILED" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
        $errors += @{Module = $module; Error = $_.Exception.Message}
    }
}

Write-Host ""

# Test New v3.0 Modules
Write-Host "Testing NEW v3.0 Modules..." -ForegroundColor Yellow
$newModules = @(
    "EmailOperations\AdvancedSearch.psm1",
    "Utilities\HealthMonitor.psm1",
    "Security\ThreatDetector.psm1",
    "Integration\CalendarSync.psm1",
    "Analytics\AttachmentStats.psm1"
)

foreach ($module in $newModules) {
    $modulePath = Join-Path $ModulesPath $module

    # Check if file exists
    if (-not (Test-Path $modulePath)) {
        Write-Host "  $module... MISSING FILE" -ForegroundColor Red
        $errors += @{Module = $module; Error = "File not found at: $modulePath"}
        continue
    }

    try {
        Write-Host "  Loading $module..." -NoNewline
        Import-Module $modulePath -Force -ErrorAction Stop -WarningVariable moduleWarnings

        if ($moduleWarnings) {
            Write-Host " OK (with warnings)" -ForegroundColor Yellow
            $warnings += @{Module = $module; Warnings = $moduleWarnings}
        } else {
            Write-Host " OK" -ForegroundColor Green
            $success += $module
        }
    } catch {
        Write-Host " FAILED" -ForegroundColor Red
        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "    Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
        $errors += @{Module = $module; Error = $_.Exception.Message; Line = $_.InvocationInfo.ScriptLineNumber}
    }
}

Write-Host ""
Write-Host "=== Summary ===" -ForegroundColor Cyan
Write-Host "Successful: $($success.Count)" -ForegroundColor Green
Write-Host "Warnings: $($warnings.Count)" -ForegroundColor Yellow
Write-Host "Errors: $($errors.Count)" -ForegroundColor Red

if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host "=== Errors Detail ===" -ForegroundColor Red
    foreach ($err in $errors) {
        Write-Host "Module: $($err.Module)" -ForegroundColor Yellow
        Write-Host "Error: $($err.Error)" -ForegroundColor Red
        if ($err.Line) {
            Write-Host "Line: $($err.Line)" -ForegroundColor Red
        }
        Write-Host ""
    }
}

if ($warnings.Count -gt 0) {
    Write-Host ""
    Write-Host "=== Warnings Detail ===" -ForegroundColor Yellow
    foreach ($warn in $warnings) {
        Write-Host "Module: $($warn.Module)" -ForegroundColor Yellow
        Write-Host "Warnings: $($warn.Warnings)" -ForegroundColor Yellow
        Write-Host ""
    }
}

Write-Host ""
Write-Host "Diagnostic complete. Press any key to exit..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
