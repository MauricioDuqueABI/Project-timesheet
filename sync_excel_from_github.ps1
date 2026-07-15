# sync_excel_from_github.ps1
# Pulls latest project_data.json from GitHub and writes all task data into the Excel timesheet.
# Scheduled to run at 5:30pm daily (after the cloud sync agent pushes at ~5:02pm).
# Preserves Priority (col 6) and Notes (col 12) — these are manual and not stored in GitHub.

$workspace    = "C:\Users\Mauricio.Duque\OneDrive - ABI LTD\Documents\Work\ProjectsManagement"
$jsonPath     = "$workspace\project_data.json"
$excelTargets = @(
    "C:\Users\Mauricio.Duque\ABI LTD\ElecENG - SoftwareDepartment\project_time_sheet.xlsx",
    "$workspace\project_time_sheet.xlsx"
)

# ─── Step 1: Pull latest data from GitHub ────────────────────────────────────
Set-Location $workspace
Write-Host "Pulling latest project_data.json from GitHub..."
git pull
if (-not $?) {
    Write-Host "WARNING: git pull failed. Proceeding with existing project_data.json."
}

if (-not (Test-Path $jsonPath)) {
    Write-Host "ERROR: project_data.json not found at $jsonPath"
    exit 1
}

$tasks = Get-Content $jsonPath -Encoding utf8 | ConvertFrom-Json
Write-Host ("Loaded " + $tasks.Count + " tasks from project_data.json")

# ─── Step 2: Open Excel ───────────────────────────────────────────────────────
$excel = New-Object -ComObject Excel.Application
$excel.Visible        = $false
$excel.DisplayAlerts  = $false
$excel.AutomationSecurity = 3   # suppress macro/security dialogs

foreach ($excelPath in $excelTargets) {

    if (-not (Test-Path $excelPath)) {
        Write-Host ("SKIP (not found): " + $excelPath)
        continue
    }

    Write-Host ("Updating: " + $excelPath)

    $wb = $excel.Workbooks.Open($excelPath, 0, $false, 5, "", "", $true, 2, "", $false, $false, 0, $false, $false, 0)
    Start-Sleep -Seconds 3

    $sheet = $null
    for ($si = 1; $si -le $wb.Sheets.Count; $si++) {
        if ($wb.Sheets.Item($si).Name -eq "KII Tasks") { $sheet = $wb.Sheets.Item($si); break }
    }
    if (-not $sheet) {
        Write-Host ("  ERROR: 'KII Tasks' sheet not found — skipping.")
        $wb.Close($false)
        continue
    }

    # Preserve manual Priority (col 6) and Notes (col 12) for existing rows
    $preserved = @{}
    $lastRow = $sheet.UsedRange.Rows.Count
    for ($r = 2; $r -le $lastRow; $r++) {
        $numVal = $sheet.Cells.Item($r, 1).Value2
        if ($null -ne $numVal -and "$numVal" -ne "") {
            $n = [int]$numVal
            $preserved[$n] = @{
                priority = "$($sheet.Cells.Item($r, 6).Value2)".Trim()
                notes    = "$($sheet.Cells.Item($r, 12).Value2)".Trim()
            }
        }
    }

    # Clear all data rows (keep header in row 1)
    if ($lastRow -gt 1) {
        $sheet.Range(
            $sheet.Cells.Item(2, 1),
            $sheet.Cells.Item($lastRow, 13)
        ).ClearContents()
    }

    # Write tasks from project_data.json
    $row = 2
    foreach ($task in $tasks) {
        $num      = $task.num
        $priority = if ($preserved.ContainsKey($num)) { $preserved[$num].priority } else { "" }
        $notes    = if ($preserved.ContainsKey($num)) { $preserved[$num].notes }    else { "" }

        $sheet.Cells.Item($row, 1).Value2  = $num
        $sheet.Cells.Item($row, 2).Value2  = $task.title
        $sheet.Cells.Item($row, 3).Value2  = $task.description
        $sheet.Cells.Item($row, 4).Value2  = $task.assignee
        $sheet.Cells.Item($row, 5).Value2  = $task.status
        $sheet.Cells.Item($row, 6).Value2  = $priority
        $sheet.Cells.Item($row, 7).Value2  = $task.size
        if ($null -ne $task.estimate -and "$($task.estimate)" -ne "") {
            $sheet.Cells.Item($row, 8).Value2 = [double]$task.estimate
        }
        if ($task.start_date -and $task.start_date -ne "") {
            $sheet.Cells.Item($row, 9).Value2  = $task.start_date
        }
        if ($task.target_date -and $task.target_date -ne "") {
            $sheet.Cells.Item($row, 10).Value2 = $task.target_date
        }
        $sheet.Cells.Item($row, 11).Value2 = $task.labels
        $sheet.Cells.Item($row, 12).Value2 = $notes
        $sheet.Cells.Item($row, 13).Value2 = $task.scope

        $row++
    }

    $wb.Save()
    $wb.Close($false)
    Write-Host ("  Done — wrote " + ($row - 2) + " tasks.")
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
Write-Host "Excel sync complete at $timestamp"
