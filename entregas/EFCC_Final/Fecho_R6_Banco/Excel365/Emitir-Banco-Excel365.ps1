param(
 [Parameter(Mandatory=$true)][string]$Source,
 [Parameter(Mandatory=$true)][string]$Recipient,
 [string]$Scenario='Proposal Bank A',
 [string]$Comparison='Official KK-01 6%',
 [string]$OutputDirectory=(Join-Path (Get-Location) ('KK01_BANK_'+(Get-Date -Format 'yyyyMMdd_HHmmss'))),
 [ValidateSet('Bank','Brief','Consolidated')][string]$Format='Bank'
)
# Optional native acceptance/emission tool. Not executed in the Linux workspace.
# Never opens/saves the user's original file for modification.
$ErrorActionPreference='Stop'
if(Test-Path $OutputDirectory){throw 'Output directory must be new: issued versions are never overwritten.'}
if([string]::IsNullOrWhiteSpace($Recipient) -or $Recipient -eq 'To identify'){throw 'Identify the report recipient.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$sourcePath=(Resolve-Path $Source).Path
$sourceHash=(Get-FileHash $sourcePath -Algorithm SHA256).Hash
$excel=$null;$book=$null;$snapshots=@();$rows=@()
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false
 # Alerts stay enabled: corruption/repair must not be silently accepted.
 $excel.DisplayAlerts=$true
 $excel.AskToUpdateLinks=$false
 $excel.AutomationSecurity=3
 foreach($profile in @($Scenario,$Comparison)) {
  $tag=if($profile -eq $Scenario){'Selected'}else{'Comparison'}
  $copy=Join-Path $OutputDirectory ($tag+'_Snapshot.xlsx')
  Copy-Item $sourcePath $copy
  $book=$excel.Workbooks.Open($copy,0,$false)
  $book.Worksheets.Item('EXEC V06').Range('H23').Value2='Test'
  $book.Worksheets.Item('DASH 03 V06').Range('H33').Value2=$profile
  $book.Worksheets.Item('DASH 10 Centro').Range('H122').Value2=$Recipient
  $excel.Calculation=-4105 # xlCalculationAutomatic
  $excel.CalculateFullRebuild()
  $deadline=(Get-Date).AddMinutes(10)
  while($excel.CalculationState -ne 0){
   if((Get-Date) -gt $deadline){throw 'Calculation did not complete. PDF emission blocked.'}
   Start-Sleep -Milliseconds 250
  }
  foreach($sheet in $book.Worksheets){
   $errorCells=$null
   try {$errorCells=$sheet.UsedRange.SpecialCells(-4123,16)}catch [System.Runtime.InteropServices.COMException]{}
   if($null -ne $errorCells){throw ('Formula error on '+$sheet.Name+'!'+$errorCells.Address())}
  }
  $motor=$book.Worksheets.Item('V06 Motor')
  for($r=5;$r -le 184;$r++) {
   foreach($c in @('DI','DJ','DZ','EJ')){
    $v=$motor.Range($c+$r).Value2
    if($null -eq $v -or [math]::Abs([double]$v) -gt 0.005){throw ('Unreconciled flow: '+$c+$r)}
   }
   $rows += [pscustomobject]@{Profile=$profile;Month=($r-4);Debt=$motor.Range('CK'+$r).Value2;Cash=$motor.Range('DG'+$r).Value2;Interest=$motor.Range('CE'+$r).Value2;Draw=$motor.Range('CD'+$r).Value2;Principal=$motor.Range('CA'+$r).Value2;EarlyPrincipal=$motor.Range('CJ'+$r).Value2;Fees=$motor.Range('EO'+$r).Value2;DebtService=$motor.Range('EP'+$r).Value2;CFADS=$motor.Range('BX'+$r).Value2}
  }
  $gate=[string]$book.Worksheets.Item('V06 Banco Servico').Range('B230').Value2
  if(-not $gate.StartsWith('CHECKS PASS')){throw ('Emission blocked: '+$gate)}
  $book.Save()
  $snapshots += [pscustomobject]@{Profile=$profile;File=$copy;NPV=$book.Worksheets.Item('V06 Cenarios').Range('M5').Value2;IRR=$book.Worksheets.Item('V06 Cenarios').Range('K5').Value2;ExcelVersion=$excel.Version;ExcelBuild=$excel.Build;SHA256=(Get-FileHash $copy -Algorithm SHA256).Hash}
  # Native prints certify the selected scenario; compare using the issued JSON/CSV.
  $printSheet=switch($Format){'Bank' {'V06 Banco Relatorio'} 'Brief' {'V06 Banco Resumo'} 'Consolidated' {'DASH 10 V06'}}
  $sheet=$book.Worksheets.Item($printSheet)
  $sheet.Visible=-1
  $sheet.ExportAsFixedFormat(0,(Join-Path $OutputDirectory ($tag+'_'+$Format+'.pdf')))
  $book.Close($false);$book=$null
  # Reopen the saved snapshot and rebuild to detect save/reopen calculation errors.
  $book=$excel.Workbooks.Open($copy,0,$true)
  $excel.CalculateFullRebuild()
  if($excel.CalculationState -ne 0){throw 'Reopened snapshot calculation incomplete.'}
  $newNPV=[double]$book.Worksheets.Item('V06 Cenarios').Range('M5').Value2
  if([math]::Abs($newNPV-[double]$snapshots[-1].NPV) -gt 0.005){throw 'NPV changed after save/reopen.'}
  $book.Close($false);$book=$null
 }
 if((Get-FileHash $sourcePath -Algorithm SHA256).Hash -ne $sourceHash){throw 'Original source changed.'}
 $rows | Export-Csv -NoTypeInformation -Encoding UTF8 (Join-Path $OutputDirectory 'Monthly_Comparison.csv')
 [pscustomobject]@{Status='Native Excel calculation checks passed';DateUTC=(Get-Date).ToUniversalTime().ToString('o');OfficialTAN=0.06;OfficialGraceMonths=42;SourceSHA256=$sourceHash;Snapshots=$snapshots;Limitation='Native export produces selected/comparison PDFs separately; Linux controlled issuer creates side-by-side Base/Test tables.'} | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 (Join-Path $OutputDirectory 'Excel365_Result.json')
} catch {
 Get-ChildItem $OutputDirectory -Filter '*.pdf' | Remove-Item
 $_.Exception.Message | Set-Content -Encoding UTF8 (Join-Path $OutputDirectory 'BLOCKED.txt')
 throw
} finally {
 if($null -ne $book){$book.Close($false)}
 if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
