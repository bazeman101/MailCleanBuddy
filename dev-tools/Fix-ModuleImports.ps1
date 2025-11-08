# Remove Import-Module statements from all .psm1 files
# The main script already imports all modules in the correct order

Write-Host "=== Fixing Module Import Statements ===" -ForegroundColor Cyan
Write-Host ""

$modulesPath = Join-Path $PSScriptRoot "Modules"
$allModules = Get-ChildItem -Path $modulesPath -Filter "*.psm1" -Recurse

$fixedCount = 0
$skippedCount = 0

foreach ($module in $allModules) {
    $content = Get-Content $module.FullName -Raw
    $originalContent = $content

    # Remove Import-Module lines (but keep comments)
    $content = $content -replace '(?m)^Import-Module\s+.*?-Force\s*$', '# [REMOVED] Import-Module statement - modules are imported by main script'

    if ($content -ne $originalContent) {
        Set-Content -Path $module.FullName -Value $content -NoNewline
        Write-Host "Fixed: $($module.FullName.Replace($modulesPath, 'Modules'))" -ForegroundColor Green
        $fixedCount++
    } else {
        Write-Host "Skipped: $($module.FullName.Replace($modulesPath, 'Modules')) (no changes needed)" -ForegroundColor Gray
        $skippedCount++
    }
}

Write-Host ""
Write-Host "=== Summary ===" -ForegroundColor Cyan
Write-Host "Fixed: $fixedCount modules" -ForegroundColor Green
Write-Host "Skipped: $skippedCount modules (no Import-Module statements)" -ForegroundColor Gray
Write-Host ""
Write-Host "All Import-Module statements have been removed from module files." -ForegroundColor Green
Write-Host "The main MailCleanBuddy.ps1 script handles all module imports." -ForegroundColor Green
