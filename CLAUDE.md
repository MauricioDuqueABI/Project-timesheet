# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this workspace is

A project management workspace that mirrors the **KII GitHub Project** (`Abi-ltd` org, project #19) into an Excel timesheet. There is no application code — the artefacts are:

- `project_time_sheet.xlsx` — the live working file (180 tasks as of 2026-06-30)
- `project_time_sheet - template.xlsx` — blank template for new projects

## GitHub data source

- **Org:** `Abi-ltd`
- **Project:** KII (project #19) — `https://github.com/orgs/Abi-ltd/projects/19`
- **Token:** stored in `~/.claude/settings.json` under `mcpServers.github.env.GITHUB_PERSONAL_ACCESS_TOKEN`
- **API:** GitHub GraphQL (`https://api.github.com/graphql`) — REST v3 cannot read ProjectsV2 field values

### Fetching all project items (paginated GraphQL)

```powershell
$token = (Get-Content "$env:USERPROFILE\.claude\settings.json" | ConvertFrom-Json).mcpServers.github.env.GITHUB_PERSONAL_ACCESS_TOKEN
$headers = @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" }

function Build-Query($afterCursor) {
    $afterPart = if ($afterCursor) { ", after: `"$afterCursor`"" } else { "" }
    '{ organization(login: "Abi-ltd") { projectV2(number: 19) { items(first: 50' + $afterPart + ') { pageInfo { hasNextPage endCursor } nodes { content { ... on Issue { number title body assignees(first: 5) { nodes { login } } labels(first: 10) { nodes { name } } } ... on DraftIssue { title body assignees(first: 5) { nodes { login } } } } fieldValues(first: 20) { nodes { ... on ProjectV2ItemFieldTextValue { text field { ... on ProjectV2Field { name } } } ... on ProjectV2ItemFieldNumberValue { number field { ... on ProjectV2Field { name } } } ... on ProjectV2ItemFieldSingleSelectValue { name field { ... on ProjectV2SingleSelectField { name } } } ... on ProjectV2ItemFieldDateValue { date field { ... on ProjectV2Field { name } } } } } } } } } }'
}
# Loop with $page.pageInfo.hasNextPage / endCursor until exhausted
```

**Important:** Build the query string via string concatenation — PowerShell here-strings mangle the cursor when paginating, causing `Problems parsing JSON` errors from the API.

## Excel file schema — "KII Tasks" sheet

| Col | Header | Source |
|-----|--------|--------|
| 1 | # | Row index |
| 2 | Title | `content.title` |
| 3 | Description | `content.body` (newlines collapsed to spaces) |
| 4 | Assignee | `content.assignees.nodes[].login` (comma-separated) |
| 5 | Status | `fieldValues` → `Status` (single-select) |
| 6 | Priority | **Inferred** (not set in GitHub) — see rules below |
| 7 | Size | `fieldValues` → `Size` (single-select: XS/S/M/L/XL) |
| 8 | Estimate (h) | `fieldValues` → `Estimate` (number) |
| 9 | Start Date | `fieldValues` → `Start date` (date, YYYY-MM-DD) |
| 10 | Target Date | `fieldValues` → `Target date` (date, YYYY-MM-DD) |
| 11 | Labels | `content.labels.nodes[].name` (comma-separated) |
| 12 | Notes | (manual) |
| 13 | Scope | `fieldValues` → `Scope` (single-select: Katana / Eye-Q / Unity / Capstone) |

Other sheets: `📋 Instructions` (user guide), `Team` (hidden, lists GitHub logins).

### GitHub project field names (exact spelling matters in switch statements)
`Status`, `Priority`, `Size`, `Estimate`, `Start date`, `Target date`, `Scope`

Note: `Start date` and `Target date` use lowercase 'd' — mismatching case means dates silently won't populate.

## Priority assignment rules

Priority is **not set** in the GitHub project for any task. It is inferred from the description using these rules (in order of precedence):

1. **Explicit tag in description** — if the body contains `**Priority:** High/Medium/Low`, honour it verbatim.
2. **High** — safety hazards (e-stop, guard stop, door-open-no-error), crashes/stability failures, security/permission gaps, operational blockers (OPC values not saved, pump pressure can't be set), core infrastructure (PLC/robot comms, RAPID program, DB bugs).
3. **Medium** — feature bugs, new features that improve workflow, core product features (EyeQ detection, stitching), integrations, testing tasks.
4. **Low** — cosmetic/UI polish, documentation, nice-to-have features, HR/admin items, investigations with no committed outcome.

Valid values: `High`, `Medium`, `Low`.

## Team (GitHub logins)

`ashiqul-abi`, `dilinaabi`, `Ksomjee`, `MauricioDuqueABI`, `mpatel0975`, `ThongHHuynhh`

## Working with Excel via PowerShell

Use the Excel COM object — do not attempt to parse `.xlsx` binary with Read/Grep tools.

```powershell
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($excelPath)
$sheet = $wb.Sheets | Where-Object { $_.Name -eq "KII Tasks" }
# ... read/write $sheet.Cells.Item($row, $col) ...
$wb.Save(); $wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
```

If `Workbooks.Open` throws "cannot access the file", a previous Excel COM session is still alive — run `Get-Process excel | Stop-Process -Force` before retrying.

Avoid `??` (null-coalescing) and `$c:` style variable references — this environment runs **Windows PowerShell 5.1**, which does not support those operators.
