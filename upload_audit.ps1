# ==========================================
# UPLOAD LATEST AUDIT JSON TO SERVER SHARE
# For Cascades — runs via ScreenConnect Toolbox
# ==========================================

$ServerShare = "\\192.168.2.254\AuditUpload"
$LocalDir = "C:\Temp"

# Find newest JSON audit file
$newest = Get-ChildItem -Path $LocalDir -Filter "*_audit_*.json" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $newest) {
    Write-Host "[FAIL] No audit JSON files found in $LocalDir" -ForegroundColor Red
    exit 1
}

Write-Host "Found: $($newest.Name) ($([math]::Round($newest.Length / 1KB, 1)) KB)"

# Build destination filename: HOSTNAME_originalname.json
$destName = "$($env:COMPUTERNAME)_$($newest.Name)"
$destPath = Join-Path $ServerShare $destName

# Test share access
if (-not (Test-Path $ServerShare)) {
    Write-Host "[FAIL] Cannot access $ServerShare — check share permissions and firewall" -ForegroundColor Red
    exit 1
}

# Copy
try {
    Copy-Item -Path $newest.FullName -Destination $destPath -Force
    Write-Host "[OK] Uploaded to $destPath" -ForegroundColor Green
} catch {
    Write-Host "[FAIL] Copy failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
