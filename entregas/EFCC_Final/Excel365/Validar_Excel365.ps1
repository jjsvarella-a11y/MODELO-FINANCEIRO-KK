<# Native Windows / Microsoft Excel Desktop acceptance test.
Run on a LOCAL copy: powershell -ExecutionPolicy Bypass -File .\Validar_Excel365.ps1
Does not modify the supplied source. Does not send data or publish anything.
#>
param([string]$WorkbookPath = (Join-Path $PSScriptRoot '..\KK01_V06_Executive_Master.xlsx'))
$ErrorActionPreference = 'Stop'
$WorkbookPath = (Resolve-Path $WorkbookPath).Path
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$resultDir = Join-Path $PSScriptRoot "Resultados_$stamp"
New-Item -ItemType Directory -Path $resultDir | Out-Null
$testPath = Join-Path $resultDir 'KK01_Teste_Nativo.xlsx'
Copy-Item -LiteralPath $WorkbookPath -Destination $testPath
$results = [System.Collections.Generic.List[object]]::new()
$excel = $null; $book = $null
function Record([string]$Name,[string]$Status,$Evidence) {
    $results.Add([pscustomobject]@{test=$Name;status=$Status;evidence=$Evidence})
    Write-Host "$Status : $Name"
}
function Recalculate {
    $excel.CalculateFullRebuild()
    $until = (Get-Date).AddMinutes(10)
    while ($excel.CalculationState -ne 0) {
        if ((Get-Date) -gt $until) { throw 'Excel calculation did not finish within ten minutes.' }
        Start-Sleep -Milliseconds 200
    }
}
function Cell($Sheet,$Address) { return $book.Worksheets.Item($Sheet).Range($Address) }
function Number($Sheet,$Address) {
    $v=(Cell $Sheet $Address).Value2
    if ($v -isnot [double] -and $v -isnot [int]) { throw "Non-numeric KPI ${Sheet}!${Address}: $v" }
    return [double]$v
}
function ErrorScan {
    $found=[System.Collections.Generic.List[object]]::new()
    $orderedSheets=@($book.Worksheets | Sort-Object @{Expression={ if ($_.Name -eq 'V06 Motor') { 0 } elseif ($_.Name -like 'V06 Calend*') { 1 } elseif ($_.Name -match '^(EXEC|DASH)') { 2 } else { 3 } }})
    foreach ($sheet in $orderedSheets) {
        foreach ($kind in @(-4123,2)) { # formulas and constants, errors only
            $bad=$null
            try { $bad=$sheet.UsedRange.SpecialCells($kind,16) } catch { $bad=$null }
            if ($null -ne $bad) {
                foreach ($cell in $bad.Cells) {
                    if ($found.Count -ge 2000) { break }
                    $found.Add([pscustomobject]@{sheet=$sheet.Name;cell=$cell.Address($false,$false);error=$cell.Text;formula=$cell.Formula})
                }
            }
        }
    }
    return ,($found.ToArray())
}
function SyncLogo {
    $identity=$book.Worksheets.Item('V06 Identidade')
    $pictures=@($identity.Shapes | Where-Object { $_.Type -eq 13 -or $_.Type -eq 11 })
    if ($pictures.Count -eq 0) { Record 'Logo insertion / header synchronisation' 'NOT TESTED' 'Insert one original JPG/PNG picture in V06 Identidade B8:H20, save and run this test again.'; return }
    if ($pictures.Count -ne 1) { Record 'Logo insertion / header synchronisation' 'FAIL' 'Keep exactly one original picture on V06 Identidade.'; return }
    $logo=$pictures[0];$count=0
    $targets=@('EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06','DASH 10 Centro','DASH 10 V06','V06 Brief','V06 PDF 01','V06 PDF 02','V06 PDF 03','V06 PDF 04','V06 PDF 05','V06 PDF 06','V06 PDF 07','V06 PDF 08','V06 PDF 09')
    foreach ($name in $targets) {
        $sheet=$book.Worksheets.Item($name)
        foreach ($shape in @($sheet.Shapes)) { if ($shape.Name -like 'KK01_JS_CV_Logo_*') { $shape.Delete() } }
        $pages=if ($name -eq 'DASH 10 V06') { 10 } elseif ($name -eq 'V06 Brief') { 2 } else { 1 }
        for ($i=0;$i -lt $pages;$i++) {
            $address=if ($name -in @('DASH 10 V06','V06 Brief')) { 'J'+(1+60*$i)+':L'+(3+60*$i) } elseif ($name -like 'V06 PDF *') { 'J1:L3' } else { 'N2:P3' }
            $area=$sheet.Range($address)
            $logo.Copy();$sheet.Activate();$sheet.Paste()
            $picture=$sheet.Shapes.Item($sheet.Shapes.Count)
            $picture.Name='KK01_JS_CV_Logo_'+$i;$picture.LockAspectRatio=-1
            $picture.Width=[math]::Max(1,$area.Width-6)
            if ($picture.Height -gt $area.Height-6) { $picture.Height=[math]::Max(1,$area.Height-6) }
            $picture.Left=$area.Left+($area.Width-$picture.Width)/2
            $picture.Top=$area.Top+($area.Height-$picture.Height)/2
            $picture.Placement=1
            $area.Cells.Item(1,1).Value2=''
            $count++
        }
    }
    $excel.CutCopyMode=$false
    Record 'Logo insertion / header synchronisation' 'PASS' @{embedded_original_count=1;header_copies=$count;financial_inputs_changed=$false;output='Test copy only'}
}
function Mutation([string]$Name,$Changes,[scriptblock]$Verify) {
    if ($Changes.Count -gt 0 -and $Changes[0] -is [string]) { $Changes=@(,$Changes) }
    $saved=@()
    try {
        foreach ($change in $Changes) {
            $cell=Cell $change[0] $change[1]
            $saved+=,@($change[0],$change[1],$cell.Formula)
            $cell.Value2=$change[2]
        }
        Recalculate
        $errors=ErrorScan
        if ($errors.Count -gt 0) { throw ($errors | ConvertTo-Json -Depth 4 -Compress) }
        $evidence=& $Verify
        Record $Name 'PASS' $evidence
    } catch { Record $Name 'FAIL' $_.Exception.Message }
    finally {
        foreach ($old in $saved) { (Cell $old[0] $old[1]).Formula=$old[2] }
        Recalculate
    }
}
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AskToUpdateLinks=$false
    $excel.AutomationSecurity=3 # msoAutomationSecurityForceDisable
    $releaseIds=$null
    try { $releaseIds=(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').ProductReleaseIds } catch {}
    $meta=[ordered]@{utc=(Get-Date).ToUniversalTime().ToString('o');excel_version=$excel.Version;product_release_ids=$releaseIds;excel_build=$excel.Build;operating_system=$excel.OperatingSystem;source=$WorkbookPath;sha256=(Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash;source_modified=$false;test_copy=$testPath}
    $meta | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 (Join-Path $resultDir 'Ambiente.json')
    if ($releaseIds -match 'O365') { Record 'Microsoft 365 product identity' 'PASS' $releaseIds } else { Record 'Microsoft 365 product identity' 'NOT TESTED' 'Version 16 alone does not identify Microsoft 365; verify installed edition.' }
    # Reproduce the reported transition: cached values in Protected View, then Edit.
    # No xlRepairFile / xlExtractData request is ever used.
    $protected=$null
    try {
        $protected=$excel.ProtectedViewWindows.Open($testPath)
        $cachedBook=$protected.Workbook
        $before=[ordered]@{}
        foreach ($item in @(@('EXEC V06','M15'),@('DASH 01 V06','I9'),@('DASH 03 V06','E9'))) {
            $before[($item[0]+'!'+$item[1])]=$cachedBook.Worksheets.Item($item[0]).Range($item[1]).Value2
        }
        $before | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 (Join-Path $resultDir 'Valores_Guardados_Antes_Edicao.json')
        Record 'Cached values before enabling editing (snapshot only)' 'PASS' $before
        $book=$protected.Edit();$protected=$null
        Record 'Protected View -> Enable Editing' 'PASS' 'Transition executed in native Excel; cached display alone is not calculation validation.'
    } catch {
        Record 'Protected View -> Enable Editing' 'NOT TESTED' $_.Exception.Message
        if ($null -ne $protected) { try { $protected.Close() } catch {} }
        $book=$excel.Workbooks.Open($testPath,0,$false)
    }
    Record 'Open for editing, no requested repair' 'PASS' @{sheets=$book.Worksheets.Count}
    $editDeadline=(Get-Date).AddMinutes(10)
    while ($excel.CalculationState -ne 0) {
        if ((Get-Date) -gt $editDeadline) { throw 'Calculation after enabling editing timed out.' }
        Start-Sleep -Milliseconds 200
    }
    $editErrors=ErrorScan
    $editErrors | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 (Join-Path $resultDir 'Erros_Apos_Ativar_Edicao.json')
    Record 'Errors immediately after enabling editing' $(if ($editErrors.Count) { 'FAIL' } else { 'PASS' }) @{error_records=$editErrors.Count;records_capped=2000}
    try { SyncLogo } catch { Record 'Logo synchronisation' 'FAIL' $_.Exception.Message }
    Recalculate
    $errors=ErrorScan
    $errors | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $resultDir 'Erros_Baseline.json')
    if ($errors.Count -gt 0) { throw "Native Excel produced $($errors.Count) error records (capped at 2000); inspect Erros_Baseline.json. Do not certify." }
    Record 'Full recalculation, all workbook errors' 'PASS' @{errors=0}
    $npv=Number 'EXEC V06' 'M15';$adr=Number 'DASH 05 V06' 'E9';$fees=Number 'DASH 05 V06' 'I15'
    $kpis=@{}
    $dashboards=@('EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06')
    foreach ($name in $dashboards) { foreach ($addr in @('E9','I9','M9','E15','I15','M15')) { $kpis["$name!$addr"]=Number $name $addr } }
    Record '54 dashboard numeric KPIs' 'PASS' $kpis
    Mutation 'PT / EN preserves financial KPIs' @(@('DASH 10 Centro','H7','EN')) {
        foreach ($name in $dashboards) { foreach ($addr in @('E9','I9','M9','E15','I15','M15')) {
            if ([math]::Abs((Number $name $addr)-$kpis["$name!$addr"]) -gt .0001) { throw 'Language changed financial result.' }
        } }
        if ((Cell 'EXEC V06' 'B8').Text -notmatch 'Executive Dashboard') { throw 'English navigation not refreshed.' }
        @{financial_kpis_unchanged=54;english_navigation=(Cell 'EXEC V06' 'B8').Text}
    }
    Mutation 'Hotel ADR propagates to investor NPV' @(@('DASH 05 V06','H23','Custom'),@('DASH 05 V06','H25',[double][math]::Min(250,$adr+5))) {
        $newadr=Number 'DASH 05 V06' 'E9';$newnpv=Number 'EXEC V06' 'M15'
        if ($newadr -eq $adr -or [math]::Abs($newnpv-$npv) -lt .01) { throw 'ADR did not propagate.' }
        @{adr=$newadr;investor_npv=$newnpv}
    }
    Mutation 'Hotel conservative active readouts' @(@('DASH 05 V06','H23','Conservative')) {
        if ([math]::Abs((Number 'DASH 05 V06' 'H24')-160) -gt .0001 -or [math]::Abs((Number 'DASH 05 V06' 'H26')-.52) -gt .000001) { throw 'Active Hotel assumptions not refreshed.' }
        @{active_adr=(Number 'DASH 05 V06' 'H24');active_occupancy=(Number 'DASH 05 V06' 'H26');stored_custom_adr=(Number 'DASH 05 V06' 'H25')}
    }
    Mutation 'Residential target active readouts' @(@('DASH 04B V06','H23','KK Target Terms')) {
        if ([math]::Abs((Number 'DASH 04B V06' 'H24')-.06) -gt .000001 -or [math]::Abs((Number 'DASH 04B V06' 'N24')-.02) -gt .000001) { throw 'Residential effective fees not refreshed.' }
        @{management=(Number 'DASH 04B V06' 'H24');distribution=(Number 'DASH 04B V06' 'N24')}
    }
    Mutation 'Equity discount active readout' @(@('DASH 07 V06','H29','8%')) {
        if ([math]::Abs((Number 'DASH 07 V06' 'N26')-.08) -gt .000001) { throw 'Effective discount not refreshed.' }
        @{discount=(Number 'DASH 07 V06' 'N26');npv=(Number 'EXEC V06' 'M15')}
    }
    $miceOld=Number 'DASH 05 V06' 'N25'
    Mutation 'MICE utilisation propagates' @(@('DASH 05 V06','N25',[double][math]::Min(.8,$miceOld+.05))) {
        $v=Number 'EXEC V06' 'M15';if ([math]::Abs($v-$npv) -lt .01) { throw 'MICE did not propagate.' };@{npv=$v}
    }
    Mutation 'Hotel optimized fees propagate' @(@('DASH 05 V06','H32','KK Optimized Terms')) {
        $v=Number 'DASH 05 V06' 'I15';if ($v -ge $fees) { throw 'Optimized fees not reduced from custom 3/3/6.' };@{operator_fees=$v;npv=(Number 'EXEC V06' 'M15')}
    }
    Mutation 'Reserve policy remains independent of dividend constraints' @(@('DASH 07 V06','H23','Conservative')) {
        $general=Number 'V06 Operacao' 'AG5';if ($general -ne 3000000) { throw 'General reserve not 3M.' };@{general=$general;div54=(Number 'DASH 07 V06' 'E9');npv=(Number 'EXEC V06' 'M15')}
    }
    Mutation 'Interest 8% propagates' @(@('DASH 03 V06','H23',[double].08)) {
        $v=Number 'EXEC V06' 'M15';if ([math]::Abs($v-$npv) -lt .01) { throw 'Interest did not propagate.' };@{npv=$v}
    }
    Mutation 'Matrix metric selector uses calculated IRR' @(@('V06 Matriz ADR','B3','IRR')) {
        $v=Number 'V06 Matriz ADR' 'C7';$expected=Number 'V06 Matriz Valores' 'D5';if ([math]::Abs($v-$expected) -gt .00000001) { throw 'Wrong IRR matrix source.' };@{irr=$v}
    }
    $gridNPV=Number 'V06 Matriz Valores' 'C5'
    Mutation 'Matrix ADR coordinate recalculates 180-month case' @(@('V06 Matriz ADR','C6',[double]155)) {
        $v=Number 'V06 Matriz Valores' 'C5';if ([math]::Abs($v-$gridNPV) -lt .01) { throw 'Matrix coordinate is static.' };@{npv=$v}
    }
    foreach ($name in @('V06 PDF 01','V06 PDF 02','V06 PDF 03','V06 PDF 04','V06 PDF 05','V06 PDF 06','V06 PDF 07','V06 PDF 08','V06 PDF 09','DASH 10 V06','V06 Brief')) {
        $sheet=$book.Worksheets.Item($name)
        $pdf=Join-Path $resultDir (($name -replace '[^a-zA-Z0-9]','_')+'.pdf')
        try { $sheet.ExportAsFixedFormat(0,$pdf,0,$true,$false);Record "Native PDF: $name" 'PASS' @{bytes=(Get-Item -LiteralPath $pdf).Length} } catch { Record "Native PDF: $name" 'FAIL' $_.Exception.Message }
    }
    $matrixPath=Join-Path $PSScriptRoot '..\Matriz_Navegacao.csv'
    $clicked=0;$unavailable=0;$routeFailures=@()
    $navigationReportFormat=(Cell 'DASH 10 Centro' 'H11').Formula
    foreach ($route in (Import-Csv -LiteralPath $matrixPath -Delimiter ';')) {
        $origin=$route.origin.Split('!');$dest=$route.destination.Split('!')
        try {
            $sheet=$book.Worksheets.Item($origin[0]);$sheet.Activate();$cell=$sheet.Range($origin[1])
            if ($origin[0] -eq 'DASH 10 Centro' -and $origin[1] -eq 'E12') {
                (Cell 'DASH 10 Centro' 'H11').Value2=$(if ($dest[0] -eq 'V06 Brief') {'Executive Brief'} else {'Complete Report'})
                $excel.Calculate()
            }
            if ($cell.Hyperlinks.Count -gt 0) {
                $cell.Hyperlinks.Item(1).Follow()
                if ($book.ActiveSheet.Name -ne $dest[0]) { throw 'Unexpected navigation destination.' }
                $clicked++
            } else { $unavailable++ }
        } catch { $routeFailures+=@{origin=$route.origin;destination=$route.destination;error=$_.Exception.Message} }
    }
    (Cell 'DASH 10 Centro' 'H11').Formula=$navigationReportFormat
    Record 'Native internal hyperlink follow' $(if ($routeFailures.Count) {'FAIL'} elseif ($unavailable) {'NOT TESTED'} else {'PASS'}) @{followed=$clicked;formula_links_without_follow_API=$unavailable;failures=$routeFailures;manual_follow_required=($unavailable -gt 0)}
    $book.Worksheets.Item('EXEC V06').Activate()
    $book.Save();$book.Close($false);$book=$null
    $book=$excel.Workbooks.Open($testPath,0,$true);Recalculate
    $errors=ErrorScan
    if ($errors.Count) { throw 'Errors after native save/reopen.' }
    Record 'Native Save / Close / Reopen / Recalculate' 'PASS' @{errors=0}
    Record 'Manual inspection: opening repair messages, logo, truncation, filters, all hyperlinks, PDF 10/2-page counts' 'NOT TESTED' 'Requires human observation in Excel Print Preview. Automatic COM opening alone does not certify every UI detail.'
} catch {
    Record 'Native acceptance execution' 'FAIL' $_.Exception.Message
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    $results | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 (Join-Path $resultDir 'Resultados.json')
    Write-Host "Evidence saved: $resultDir"
}
