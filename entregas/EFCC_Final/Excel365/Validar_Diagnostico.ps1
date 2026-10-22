param([string]$WorkbookPath=(Join-Path $PSScriptRoot 'Diagnostico_Formulas_Excel365.xlsx'))
$ErrorActionPreference='Stop'
$folder=Join-Path $PSScriptRoot ('Diagnostico_'+(Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $folder | Out-Null
$copy=Join-Path $folder 'Diagnostico_Calculado_Excel.xlsx'
Copy-Item -LiteralPath $WorkbookPath -Destination $copy
$excel=$null;$book=$null;$protected=$null
$result=[ordered]@{native_run=$false;protected_view_edit='NOT TESTED';cells=@();errors=@();exception=$null}
try {
    $excel=New-Object -ComObject Excel.Application
    $result.native_run=$true
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3
    $result.version=$excel.Version;$result.build=$excel.Build
    try {
        $protected=$excel.ProtectedViewWindows.Open($copy)
        $result.cached_before_edit=$protected.Workbook.Worksheets.Item('DIAGNOSTICO').Range('B5').Value2
        $book=$protected.Edit();$protected=$null
        $result.protected_view_edit='PASS'
    } catch {
        $result.protected_view_edit_exception=$_.Exception.Message
        if ($null -ne $protected) { try { $protected.Close() } catch {} }
        $book=$excel.Workbooks.Open($copy,0,$false)
    }
    $excel.CalculateFullRebuild()
    $deadline=(Get-Date).AddMinutes(3)
    while ($excel.CalculationState -ne 0) {
        if ((Get-Date) -gt $deadline) { throw 'Native calculation timed out.' }
        Start-Sleep -Milliseconds 200
    }
    foreach ($entry in @(@('V06 Motor','G5'),@('V06 Motor','H5'),@('V06 Motor','O5'),@('V06 Motor','P5'),@('V06 Curva','A5'),@('V06 Curva','B5'),@('V06 Motor','A5'),@('V06 Motor','C5'),@('V06 Unidades','H5'),@('V06 Unidades','I5'),@('V06 Unidades','L5'),@('DIAGNOSTICO','B5'),@('DIAGNOSTICO','C5'),@('DIAGNOSTICO','B6'),@('DIAGNOSTICO','C6'))) {
        $cell=$book.Worksheets.Item($entry[0]).Range($entry[1]);$formula2=$null
        try { $formula2=$cell.Formula2 } catch {}
        $v=$cell.Value2
        $result.cells+=,[pscustomobject]@{sheet=$entry[0];cell=$entry[1];text=$cell.Text;value=$v;value_type=$(if ($null -eq $v) { 'null' } else { $v.GetType().FullName });formula=$cell.Formula;formula2=$formula2}
    }
    foreach ($sheet in $book.Worksheets) {
        $bad=$null
        try { $bad=$sheet.UsedRange.SpecialCells(-4123,16) } catch {}
        if ($null -ne $bad) {
            foreach ($cell in $bad.Cells) { $result.errors+=,[pscustomobject]@{sheet=$sheet.Name;cell=$cell.Address($false,$false);text=$cell.Text;formula=$cell.Formula} }
        }
    }
    $book.Save()
    $result.status=if ($result.errors.Count -eq 0) { 'PASS' } else { 'FAIL' }
} catch { $result.status='FAIL';$result.exception=$_.Exception.Message }
finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    $result | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 (Join-Path $folder 'Resultado_Diagnostico.json')
    Write-Host "Evidence: $folder"
}
