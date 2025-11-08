# Parser Error Finder - Pinpoint exact line of syntax errors
# Usage: .\Find-ParserError.ps1 -FilePath ".\Modules\EmailOperations\EmailActions.psm1"

param(
    [Parameter(Mandatory = $false)]
    [string]$FilePath = ".\Modules\EmailOperations\EmailActions.psm1"
)

Write-Host "=== Parser Error Finder ===" -ForegroundColor Cyan
Write-Host "Analyzing: $FilePath" -ForegroundColor Yellow
Write-Host ""

if (-not (Test-Path $FilePath)) {
    Write-Host "ERROR: File not found: $FilePath" -ForegroundColor Red
    exit 1
}

# Read the entire file
$content = Get-Content $FilePath -Raw
$lines = Get-Content $FilePath

Write-Host "Testing full file parse..." -ForegroundColor Cyan

# Test full file parse
$errors = $null
$tokens = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)

if ($errors -and $errors.Count -gt 0) {
    Write-Host "Found $($errors.Count) parser error(s)!" -ForegroundColor Red
    Write-Host ""

    foreach ($error in $errors) {
        Write-Host "=== ERROR DETAILS ===" -ForegroundColor Red
        Write-Host "Message: $($error.Message)" -ForegroundColor Yellow

        if ($error.Token) {
            $lineNum = $error.Token.StartLine
            $colNum = $error.Token.StartColumn

            Write-Host "Line: $lineNum" -ForegroundColor Yellow
            Write-Host "Column: $colNum" -ForegroundColor Yellow

            # Show the problematic line and surrounding context
            Write-Host ""
            Write-Host "Context:" -ForegroundColor Cyan

            $startLine = [Math]::Max(0, $lineNum - 3)
            $endLine = [Math]::Min($lines.Count - 1, $lineNum + 2)

            for ($i = $startLine; $i -le $endLine; $i++) {
                $lineContent = $lines[$i]
                $lineNumber = $i + 1

                if ($lineNumber -eq $lineNum) {
                    Write-Host ">>> $lineNumber : $lineContent" -ForegroundColor Red
                    # Show column pointer
                    $pointer = " " * ($colNum + 6 + ([string]$lineNumber).Length) + "^"
                    Write-Host $pointer -ForegroundColor Red
                } else {
                    Write-Host "    $lineNumber : $lineContent" -ForegroundColor Gray
                }
            }
        }
        Write-Host ""
    }

    # Search for common patterns that cause this specific error
    Write-Host "=== COMMON PROBLEMATIC PATTERNS ===" -ForegroundColor Cyan

    # Pattern 1: Double quotes with time formats
    Write-Host "`nSearching for time format strings in double quotes..." -ForegroundColor Yellow
    $timePatternMatches = Select-String -Path $FilePath -Pattern '"[^"]*HH:mm[^"]*"' -AllMatches
    if ($timePatternMatches) {
        foreach ($match in $timePatternMatches) {
            Write-Host "  Line $($match.LineNumber): $($match.Line.Trim())" -ForegroundColor Magenta
        }
    }

    # Pattern 2: @ symbol in double quotes
    Write-Host "`nSearching for '@' in double-quoted strings..." -ForegroundColor Yellow
    $atPatternMatches = Select-String -Path $FilePath -Pattern '"[^"]*''@[^"]*"' -AllMatches
    if ($atPatternMatches) {
        foreach ($match in $atPatternMatches) {
            Write-Host "  Line $($match.LineNumber): $($match.Line.Trim())" -ForegroundColor Magenta
        }
    }

    # Pattern 3: Double quotes with $var: pattern
    Write-Host "`nSearching for variable:text pattern in double quotes..." -ForegroundColor Yellow
    $varColonMatches = Select-String -Path $FilePath -Pattern '"[^"]*\$\w+:[^"]*"' -AllMatches
    if ($varColonMatches) {
        foreach ($match in $varColonMatches) {
            Write-Host "  Line $($match.LineNumber): $($match.Line.Trim())" -ForegroundColor Magenta
        }
    }

    # Pattern 4: Any double-quoted string with colon after variable-like pattern
    Write-Host "`nSearching for any ':' that might be interpreted as variable..." -ForegroundColor Yellow
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        # Look for patterns like "text: or " : or similar in double quotes
        if ($line -match '"[^"]*\w+:[^"]*"' -and $line -notmatch "'[^']*\w+:[^']*'") {
            Write-Host "  Line $($i + 1): $($line.Trim())" -ForegroundColor Magenta
        }
    }

} else {
    Write-Host "No parser errors found! File syntax is valid." -ForegroundColor Green
}

Write-Host ""
Write-Host "Analysis complete." -ForegroundColor Cyan
