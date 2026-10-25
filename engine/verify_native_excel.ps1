param([Parameter(Mandatory=$true)][string]$RequestPath)
$ErrorActionPreference = 'Stop'
$request = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$sourcePath=[IO.Path]::GetFullPath($request.source_path)
$copyPath=[IO.Path]::GetFullPath($request.copy_path)
if ($sourcePath.Equals($copyPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'Native verification must open a new disposable copy, never its source.' }
if ($copyPath.Length -gt 218) { throw 'Native Excel verification copy exceeds the 218-character operating limit. Use a shorter output directory.' }
if (-not (Test-Path -LiteralPath $copyPath -PathType Leaf)) { throw 'Native verification copy does not exist.' }
$excel = $null
$workbook = $null
$bootstrap = $null
$ownedExcel = $false
function Recalculate-Checked {
    if ($request.mode -eq 'WORKBOOK') {
        $excel.CalculateFull()
    } else {
        # The edit engine emits dependency order. Two passes also tolerate a
        # summary-first check list without refreshing unrelated external links.
        for ($pass=0; $pass -lt 2; $pass++) {
            foreach ($check in $request.checks) {
                $cell = $workbook.Worksheets.Item($check.sheet).Range($check.cell)
                if ($cell.HasFormula) { $cell.Calculate() }
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
            }
        }
    }
}
function Read-Checks($Checks) {
    $rows = @()
    foreach ($check in $Checks) {
        $cell = $workbook.Worksheets.Item($check.sheet).Range($check.cell)
        $rows += [ordered]@{sheet=$check.sheet; cell=$check.cell; value=$cell.Value2; text=$cell.Text; formula=$cell.Formula}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
    }
    return ,$rows
}
try {
    $excel = New-Object -ComObject Excel.Application
    if ($excel.Workbooks.Count -ne 0) { throw 'A newly isolated Excel instance is required; existing user documents will not be touched.' }
    $ownedExcel = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $bootstrap = $excel.Workbooks.Add()
    $excel.Calculation = -4135
    $workbook = $excel.Workbooks.Open($copyPath, 0, $false)
    if ($workbook.ReadOnly) { throw 'Verification copy opened read-only.' }
    Recalculate-Checked
    $baseline = Read-Checks $request.checks
    $mutationResults = @()
    foreach ($mutation in $request.mutations) {
        $originalInputs = @()
        foreach ($inputCell in $mutation.inputs) {
            $cell = $workbook.Worksheets.Item($inputCell.sheet).Range($inputCell.cell)
            if ($cell.HasFormula) { throw 'Mutation target is a formula cell.' }
            $originalInputs += [ordered]@{sheet=$inputCell.sheet; cell=$inputCell.cell; value=$cell.Value2}
            if ($null -eq $inputCell.value) { $cell.ClearContents() }
            elseif ($inputCell.value -is [ValueType]) { $cell.Value2 = [double]$inputCell.value }
            else { $cell.Value2 = [string]$inputCell.value }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
        }
        Recalculate-Checked
        $mutationResults += [ordered]@{name=$mutation.name; inputs=$mutation.inputs; checks=(Read-Checks $mutation.checks)}
        foreach ($inputCell in $originalInputs) {
            $cell = $workbook.Worksheets.Item($inputCell.sheet).Range($inputCell.cell)
            if ($null -eq $inputCell.value) { $cell.ClearContents() }
            elseif ($inputCell.value -is [ValueType]) { $cell.Value2 = [double]$inputCell.value }
            else { $cell.Value2 = [string]$inputCell.value }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
        }
        Recalculate-Checked
    }
    $restored = Read-Checks $request.checks
    $workbook.Save()
    $report = [ordered]@{version=[string]$excel.Version; build=[string]$excel.Build; mode=$request.mode; copy_path=$request.copy_path;
        update_links=0; macros='FORCE_DISABLE'; calculation='MANUAL'; baseline=$baseline; mutations=$mutationResults; restored=$restored}
    [System.IO.File]::WriteAllText($request.output, ($report | ConvertTo-Json -Depth 30), [System.Text.UTF8Encoding]::new($false))
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $bootstrap) {
        $bootstrap.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($bootstrap)
    }
    if ($null -ne $excel -and $ownedExcel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
