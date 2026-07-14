# sync_prepare.ps1
# Run from the ProjectsManagement folder.
# 1. Copies Excel from Teams folder (replace existing)
# 2. Exports all task rows to project_data.json via Excel COM
# 3. Commits and pushes both files to GitHub

$workspace  = "C:\Users\Mauricio.Duque\OneDrive - ABI LTD\Documents\Work\ProjectsManagement"
$source     = "C:\Users\Mauricio.Duque\ABI LTD\ElecENG - SoftwareDepartment\project_time_sheet.xlsx"
$excelPath  = "$workspace\project_time_sheet.xlsx"
$jsonPath   = "$workspace\project_data.json"

# ─── Step 1: Copy ─────────────────────────────────────────────────────────────
if (Test-Path $source) {
    Copy-Item $source $excelPath -Force
    Write-Host "Copied Excel from Teams folder."
} else {
    Write-Host "WARNING: Source not found at: $source"
    Write-Host "Proceeding with existing file at destination."
}

# ─── Step 2: Export Excel → project_data.json ──────────────────────────────────
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3  # disable macro warnings that block COM calls
$wb    = $excel.Workbooks.Open($excelPath, 0, $true, 5, "", "", $true, 2, "", $false, $false, 0, $false, $false, 0)
Start-Sleep -Seconds 3  # wait for workbook to fully load before enumerating sheets

$sheet = $null
for ($si = 1; $si -le $wb.Sheets.Count; $si++) {
    if ($wb.Sheets.Item($si).Name -eq "KII Tasks") { $sheet = $wb.Sheets.Item($si); break }
}
if (-not $sheet) {
    Write-Host "ERROR: Sheet 'KII Tasks' not found."
    $wb.Close($false); $excel.Quit()
    exit 1
}

function ConvertDate($v) {
    if ($null -eq $v -or "$v".Trim() -eq "") { return "" }
    if ($v -is [double] -or $v -is [float] -or $v -is [int]) {
        return [DateTime]::FromOADate($v).ToString("yyyy-MM-dd")
    }
    return "$v".Trim()
}

$lastRow = $sheet.UsedRange.Rows.Count
$tasks = @()
for ($row = 2; $row -le $lastRow; $row++) {
    $numVal = $sheet.Cells.Item($row, 1).Value2
    if ($null -eq $numVal -or "$numVal" -eq "") { continue }

    $estVal = $sheet.Cells.Item($row, 8).Value2

    $tasks += [PSCustomObject]@{
        num         = [int]$numVal
        title       = "$($sheet.Cells.Item($row, 2).Value2)".Trim()
        description = ("$($sheet.Cells.Item($row, 3).Value2)" -replace "`r`n|`n|`r", " ").Trim()
        assignee    = "$($sheet.Cells.Item($row, 4).Value2)".Trim()
        status      = "$($sheet.Cells.Item($row, 5).Value2)".Trim()
        priority    = "$($sheet.Cells.Item($row, 6).Value2)".Trim()
        size        = "$($sheet.Cells.Item($row, 7).Value2)".Trim()
        estimate    = if ($null -ne $estVal) { $estVal } else { $null }
        start_date  = ConvertDate($sheet.Cells.Item($row, 9).Value2)
        target_date = ConvertDate($sheet.Cells.Item($row, 10).Value2)
        labels      = "$($sheet.Cells.Item($row, 11).Value2)".Trim()
        notes       = "$($sheet.Cells.Item($row, 12).Value2)".Trim()
        scope       = "$($sheet.Cells.Item($row, 13).Value2)".Trim()
    }
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$tasks | ConvertTo-Json -Depth 3 | Out-File $jsonPath -Encoding utf8
Write-Host "Exported $($tasks.Count) tasks to project_data.json"

# ─── Step 3: Git commit and push ───────────────────────────────────────────────
Set-Location $workspace
git add "project_time_sheet.xlsx" "project_data.json"
if ($?) {
    $dateStr = Get-Date -Format "yyyy-MM-dd HH:mm"
    git commit -m "Sync $dateStr"
    if ($?) {
        git push
        if ($?) {
            Write-Host "Pushed to GitHub successfully."
            Write-Host ""
            Write-Host "Next: fire the Claude sync trigger to update GitHub Projects and Monday.com."
        } else {
            Write-Host "ERROR: git push failed. Check your credentials and network."
        }
    } else {
        Write-Host "Nothing new to commit (file unchanged since last sync)."
    }
} else {
    Write-Host "ERROR: git add failed."
}
