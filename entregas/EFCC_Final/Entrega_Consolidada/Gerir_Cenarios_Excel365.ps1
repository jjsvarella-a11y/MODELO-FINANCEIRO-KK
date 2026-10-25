<# Native Microsoft Excel Desktop versioning, without VBA or Python.
   Close other copies of this workbook before running.
   .\Gerir_Cenarios_Excel365.ps1 -Action SaveBase
   .\Gerir_Cenarios_Excel365.ps1 -Action SaveComparative
   .\Gerir_Cenarios_Excel365.ps1 -Action Restore -VersionID BASE-...
   This runner was not executed in the Linux delivery environment.
#>
param(
 [ValidateSet('SaveBase','SaveComparative','Restore','Validate')][string]$Action='Validate',
 [string]$VersionID='',
 [string]$WorkbookPath=(Join-Path $PSScriptRoot 'KK01_Executive_Financial_Control_Center.xlsx')
)
$ErrorActionPreference='Stop'
$map=Get-Content (Join-Path $PSScriptRoot 'Mapa_Cenarios.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$history=Join-Path $PSScriptRoot 'Historico'
New-Item -ItemType Directory -Force -Path $history | Out-Null
$excel=$null;$book=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false
 $book=$excel.Workbooks.Open([IO.Path]::GetFullPath($WorkbookPath),0,$false)
 if ($book.ReadOnly) { throw 'Workbook is read-only. Close the open workbook and retry.' }
 $mode=if ($Action -eq 'SaveComparative') {'Comparative'} else {'Base'}
 if ($Action -eq 'Restore') {
  if ([string]::IsNullOrWhiteSpace($VersionID)) { $VersionID=[string]$book.Worksheets.Item('V07 Estado').Range('B19').Value2 }
  if ($VersionID -notmatch '^(BASE|CMP)-[A-Za-z0-9]+$') { throw 'Invalid version ID.' }
  $snapshot=Get-Content (Join-Path $history ($VersionID+'.json')) -Raw -Encoding UTF8 | ConvertFrom-Json
  $mode=$snapshot.classification
  foreach($d in $map.definitions) {
   $raw=$snapshot.raw_inputs.($d.key)
   if ($null -eq $raw) { throw ('Missing snapshot input '+$d.key) }
   $sheet=if ($mode -eq 'Base') {'V07 Base'} else {'V07 Comparativo'}
   $panel=if ($mode -eq 'Base') {$d.base_panel} else {$d.comparative_panel}
   if (-not $d.calculated_alias) {
    if ($null -ne $panel) { $book.Worksheets.Item($panel.sheet).Range($panel.cell).Value2=$raw.selection }
    else { $book.Worksheets.Item($sheet).Range('B'+$d.row).Value2=$raw.selection }
   }
   $book.Worksheets.Item($sheet).Range('C'+$d.row).Value2=$raw.custom
  }
 }
 $excel.CalculateFullRebuild()
 while ($excel.CalculationState -ne 0) { Start-Sleep -Milliseconds 200 }
 # Native formula-error scan over all worksheets, not only visible dashboards.
 foreach($ws in $book.Worksheets) {
  $errRange=$null
  try { $errRange=$ws.UsedRange.SpecialCells(-4123,16) } catch {}
  if ($null -ne $errRange) { throw ('Formula error: '+$ws.Name+'!'+$errRange.Address()) }
 }
 foreach($prefix in @('V07 B Motor','V07 C Motor')) {
  $ws=$book.Worksheets.Item($prefix)
  foreach($column in @('DI','DJ','DZ','EJ')) {
   $range=$ws.Range($column+'5:'+$column+'184').Value2
   for($i=1;$i -le 180;$i++) {
    if ([Math]::Abs([double]$range[$i,1]) -gt 0.005) { throw ('Reconciliation failed: '+$prefix+'!'+$column+($i+4)) }
   }
  }
 }
 if ($Action -ne 'Validate') {
  $sheetName=if($mode -eq 'Base') {'V07 Base'} else {'V07 Comparativo'}
  $sheet=$book.Worksheets.Item($sheetName);$state=$book.Worksheets.Item('V07 Estado')
  $idCell=if($mode -eq 'Base') {'B5'} else {'B6'}
  $previous=[string]$state.Range($idCell).Value2
  if (-not (Test-Path (Join-Path $history ($previous+'.json')))) { throw ('Previous version snapshot missing: '+$previous+'. Retain the Historico directory supplied with the workbook.') }
  $id=($(if($mode -eq 'Base'){'BASE-'}else{'CMP-'})+[DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffffffZ'))
  $raw=@{};$effective=@{};$changes=@()
  foreach($d in $map.definitions) {
   $panel=if($mode -eq 'Base'){$d.base_panel}else{$d.comparative_panel}
   $selected=if($null -ne $panel){$book.Worksheets.Item($panel.sheet).Range($panel.cell).Value2}else{$sheet.Range('B'+$d.row).Value2}
   $custom=$sheet.Range('C'+$d.row).Value2;$new=$sheet.Range('D'+$d.row).Value2;$old=$sheet.Range('E'+$d.row).Value2
   $raw[$d.key]=@{selection=$selected;custom=$custom};$effective[$d.key]=$new
   if ($new -ne $old) { $changes+=@{key=$d.key;before=$old;after=$new} }
   $oldSelection=$sheet.Range('J'+$d.row).Value2;$oldCustom=$sheet.Range('K'+$d.row).Value2
   if ($selected -ne $oldSelection) { $changes+=@{key=($d.key+'.selection');before=$oldSelection;after=$selected} }
   if ($custom -ne $oldCustom) { $changes+=@{key=($d.key+'.custom');before=$oldCustom;after=$custom} }
  }
  $metrics=@{};$mc=if($mode -eq 'Base'){'B'}else{'C'};$results=$book.Worksheets.Item('V07 Resultados')
  for($r=6;$r -le 19;$r++) { $metrics[[string]$results.Range('A'+$r).Value2]=$results.Range($mc+$r).Value2 }
  $engine=$book.Worksheets.Item($(if($mode -eq 'Base'){'V07 B Motor'}else{'V07 C Motor'}))
  $grid=$engine.Range('A5:EQ184').Value2;$monthly=@()
  for($r=1;$r -le 180;$r++) { $row=@();for($c=1;$c -le 147;$c++) {$row+=,$grid[$r,$c]};$monthly+=,@($row) }
  $record=@{id=$id;previous_id=$previous;classification=$mode;saved_at=[DateTime]::UtcNow.ToString('o');raw_inputs=$raw;effective_inputs=$effective;metrics=$metrics;monthly=$monthly;native_excel365='Executed by Excel Desktop COM on user machine';excel_version=$excel.Version}
  $json=$record | ConvertTo-Json -Depth 30
  [IO.File]::WriteAllText((Join-Path $history ($id+'.json')),$json,[Text.UTF8Encoding]::new($false))
  $log=$book.Worksheets.Item('V07 Historico');$last=$log.Cells.Item($log.Rows.Count,1).End(-4162).Row
  foreach($change in $changes) {
   $last++;$vals=@($id,$mode,$record.saved_at,$change.key,$change.before,$change.after,$metrics['Investor NPV'],$metrics['Investor IRR'],('Native Excel save; previous '+$previous),'')
   for($c=0;$c -lt $vals.Count;$c++) { $log.Cells.Item($last,$c+1).Value2=$vals[$c] }
  }
  foreach($d in $map.definitions) { $sheet.Range('E'+$d.row).Value2=$effective[$d.key];$sheet.Range('J'+$d.row).Value2=$raw[$d.key].selection;$sheet.Range('K'+$d.row).Value2=$raw[$d.key].custom }
  $state.Range($idCell).Value2=$id;$state.Range('B8').Value2=$record.saved_at
  $savedRow=$state.Cells.Item($state.Rows.Count,4).End(-4162).Row+1;$state.Cells.Item($savedRow,4).Value2=$id
  $excel.CalculateFullRebuild()
  $book.Save()
  Write-Output ('Saved '+$id+' with native Excel recalculation and 180-month checks.')
 } else {
  $book.Save();Write-Output 'Native Excel recalculation and reconciliation passed.'
 }
} finally {
 if ($null -ne $book) { $book.Close($false) }
 if ($null -ne $excel) { $excel.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
