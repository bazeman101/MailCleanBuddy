# Test ALL modules for parser errors
# This will show which modules have syntax issues

Write-Host "=== Testing ALL Modules for Parser Errors ===" -ForegroundColor Cyan
Write-Host ""

$allModules = Get-ChildItem -Path ".\Modules" -Filter "*.psm1" -Recurse | Sort-Object FullName
$totalModules = $allModules.Count
$errorModules = @()
$cleanModules = @()

Write-Host "Found $totalModules modules to test`n" -ForegroundColor Yellow

foreach ($module in $allModules) {
    $relativePath = $module.FullName.Replace((Get-Location).Path, ".").Replace("\", "/")
    Write-Host "Testing: $relativePath" -NoNewline

    try {
        $content = Get-Content $module.FullName -Raw
        $errors = $null
        $tokens = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)

        if ($errors -and $errors.Count -gt 0) {
            Write-Host " FAILED ($($errors.Count) error(s))" -ForegroundColor Red
            $errorModules += [PSCustomObject]@{
                Module = $relativePath
                ErrorCount = $errors.Count
                Errors = $errors
            }
        } else {
            Write-Host " OK" -ForegroundColor Green
            $cleanModules += $relativePath
        }
    } catch {
        Write-Host " EXCEPTION: $($_.Exception.Message)" -ForegroundColor Red
        $errorModules += [PSCustomObject]@{
            Module = $relativePath
            ErrorCount = 1
            Errors = @([PSCustomObject]@{Message = $_.Exception.Message})
        }
    }
}

Write-Host ""
Write-Host "=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total Modules: $totalModules" -ForegroundColor White
Write-Host "Clean Modules: $($cleanModules.Count)" -ForegroundColor Green
Write-Host "Modules with Errors: $($errorModules.Count)" -ForegroundColor Red

if ($errorModules.Count -gt 0) {
    Write-Host ""
    Write-Host "=== MODULES WITH ERRORS ===" -ForegroundColor Red
    foreach ($errMod in $errorModules) {
        Write-Host ""
        Write-Host "Module: $($errMod.Module)" -ForegroundColor Yellow
        Write-Host "Errors: $($errMod.ErrorCount)" -ForegroundColor Red

        foreach ($err in $errMod.Errors) {
            Write-Host "  - $($err.Message)" -ForegroundColor Red
            if ($err.Token) {
                Write-Host "    Line: $($err.Token.StartLine), Column: $($err.Token.StartColumn)" -ForegroundColor Yellow
            }
        }
    }

    Write-Host ""
    Write-Host "Run Find-ParserError.ps1 on each failing module for details" -ForegroundColor Yellow
} else {
    Write-Host ""
    Write-Host "All modules are clean! No parser errors found." -ForegroundColor Green
}

Write-Host ""
