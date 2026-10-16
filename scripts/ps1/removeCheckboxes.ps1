param (
    [string]$Path = "."
)

Write-Host "Scanning for Markdown files in: $(Resolve-Path $Path)" -ForegroundColor Cyan

# Find all markdown files recursively
$markdownFiles = Get-ChildItem -Path $Path -Filter "*.md" -Recurse

foreach ($file in $markdownFiles) {
    Write-Host "Processing: $($file.FullName)" -ForegroundColor Yellow
    
    # Read the file content as a single string with UTF-8 encoding
    $content = Get-Content -Path $file.FullName -Raw -Encoding utf8
    
    if (-not $content) { continue }

    # Regex pattern: matches lines starting with optional whitespace, a hyphen, 
    # and brackets containing a space, x, X, or nothing (e.g., - [ ], - [x], - [])
    $pattern = '(?m)^(\s*-\s*)\[[ xX]?\]\s*'
    
    # Replace the checkbox part, keeping the leading whitespace and hyphen ($1)
    $newContent = [System.Text.RegularExpressions.Regex]::Replace($content, $pattern, '$1')
    
    # Save back to the file only if changes were made
    if ($content -ne $newContent) {
        Set-Content -Path $file.FullName -Value $newContent -Encoding utf8
        Write-Host "  -> Cleaned checkboxes in $($file.Name)" -ForegroundColor Green
    } else {
        Write-Host "  -> No checkboxes found in $($file.Name)" -ForegroundColor Gray
    }
}

Write-Host "Finished processing all markdown files!" -ForegroundColor Cyan