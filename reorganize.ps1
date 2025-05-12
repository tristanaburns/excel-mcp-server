# This script ensures all directories and files are properly organized
# in the Excel MCP Server project

Write-Host "Finalizing Excel MCP Server directory reorganization..." -ForegroundColor Cyan

# Create any missing directories
$directories = @(
    "build",
    "docs/api",
    "docs/deployment",
    "docs/features",
    "health",
    "tests"
)

foreach ($dir in $directories) {
    $path = Join-Path -Path "." -ChildPath $dir
    if (-not (Test-Path -Path $path)) {
        Write-Host "Creating directory: $dir" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

# Move any remaining files to their proper locations
# Check for health files
if (Test-Path -Path "./health.py") {
    Write-Host "Moving health.py to health directory" -ForegroundColor Yellow
    Move-Item -Path "./health.py" -Destination "./health/" -Force
}

if (Test-Path -Path "./health.js") {
    Write-Host "Moving health.js to health directory" -ForegroundColor Yellow
    Move-Item -Path "./health.js" -Destination "./health/" -Force
}

# Check for build scripts
$buildScripts = @(
    "build.ps1", "build.sh",
    "test.ps1", "test.sh",
    "lint.ps1", "lint.sh",
    "deploy.ps1", "deploy.sh",
    "docker-entrypoint.sh"
)

foreach ($script in $buildScripts) {
    if (Test-Path -Path "./$script") {
        Write-Host "Moving $script to build directory" -ForegroundColor Yellow
        Move-Item -Path "./$script" -Destination "./build/" -Force
    }
}

# Update Docker related file permissions if needed (applies only when running in a Unix/Linux environment)
if (Test-Path -Path "./build/docker-entrypoint.sh") {
    Write-Host "Updating docker-entrypoint.sh permissions" -ForegroundColor Yellow
    # This will work in Linux but silently fail in Windows, which is ok
    try {
        & chmod +x ./build/docker-entrypoint.sh 2>$null
    } catch {}
}

# Update documentation references in any markdown files we may have missed
$docsFiles = Get-ChildItem -Path "." -Filter "*.md" -Recurse
foreach ($file in $docsFiles) {
    $content = Get-Content -Path $file.FullName -Raw
    
    # Skip files in the docs directory, they're already updated
    if ($file.FullName -like "*/docs/*" -or $file.FullName -like "*\docs\*") {
        continue
    }
    
    $modified = $false
    
    # Update common documentation references
    if ($content -match "\[TOOLS\.md\]\(TOOLS\.md\)") {
        $content = $content -replace "\[TOOLS\.md\]\(TOOLS\.md\)", "[TOOLS.md](docs/api/TOOLS.md)"
        $modified = $true
    }
    
    if ($content -match "\[API\.md\]\(API\.md\)") {
        $content = $content -replace "\[API\.md\]\(API\.md\)", "[API.md](docs/api/API.md)"
        $modified = $true
    }
    
    if ($content -match "\[HEALTH_CHECKS\.md\]\(HEALTH_CHECKS\.md\)") {
        $content = $content -replace "\[HEALTH_CHECKS\.md\]\(HEALTH_CHECKS\.md\)", "[HEALTH_CHECKS.md](docs/deployment/HEALTH_CHECKS.md)"
        $modified = $true
    }
    
    if ($content -match "\[BUILD_AND_DEPLOYMENT\.md\]\(BUILD_AND_DEPLOYMENT\.md\)") {
        $content = $content -replace "\[BUILD_AND_DEPLOYMENT\.md\]\(BUILD_AND_DEPLOYMENT\.md\)", "[BUILD_AND_DEPLOYMENT.md](docs/deployment/BUILD_AND_DEPLOYMENT.md)"
        $modified = $true
    }
    
    if ($content -match "\[FILE_OPERATIONS\.md\]\(FILE_OPERATIONS\.md\)") {
        $content = $content -replace "\[FILE_OPERATIONS\.md\]\(FILE_OPERATIONS\.md\)", "[FILE_OPERATIONS.md](docs/features/FILE_OPERATIONS.md)"
        $modified = $true
    }
    
    if ($content -match "\[CSV_OPERATIONS\.md\]\(CSV_OPERATIONS\.md\)") {
        $content = $content -replace "\[CSV_OPERATIONS\.md\]\(CSV_OPERATIONS\.md\)", "[CSV_OPERATIONS.md](docs/features/CSV_OPERATIONS.md)"
        $modified = $true
    }
    
    # Save changes if any were made
    if ($modified) {
        Write-Host "Updating documentation references in $($file.Name)" -ForegroundColor Yellow
        Set-Content -Path $file.FullName -Value $content
    }
}

Write-Host "Directory reorganization complete!" -ForegroundColor Green
