Install-Module -Name ImportExcel -Scope CurrentUser
Import-Module ImportExcel
set-executionpolicy remotesigned
$Excelobj = New-Object -ComObject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open('C:\Imporfile\ExcelTraining.xlsx')
$filePath = "C:\Imporfile\ExcelTraining.xlsx"
$excelData = Import-Excel -Path $filePath -Noheader
$Excelworksheet = $ExcelWorkBook.sheets.Item("ServerList")
$name = $ExcelWorksheet.columns.Item(1).Rows.Item(100).Text
$funcAbbr  = $ExcelWorksheet.columns.Item(2).Rows.Item(100).Text
$serverDesig = $ExcelWorksheet.columns.Item(3).Rows.Item(100).Text
$windowsOrlinux = $ExcelWorksheet.columns.Item(4).Rows.Item(100).Text
$envi = $ExcelWorksheet.columns.Item(5).Rows.Item(100).Text
$exclude = $ExcelWorksheet.columns.Item(6).Rows.Item(100).Text
$appName = $ExcelWorksheet.columns.Item(7).Rows.Item(100).Text
$serverType = $ExcelWorksheet.columns.Item(8).Rows.Item(100).Text
$buildOrder = $ExcelWorksheet.columns.Item(9).Rows.Item(100).Text
$pbmSolution = $ExcelWorksheet.columns.Item(10).Rows.Item(100).Text
$cefSolution = $ExcelWorksheet.columns.Item(11).Rows.Item(100).Text
$instancetype = $ExcelWorksheet.columns.Item(12).Rows.Item(100).Text
$estMonthlyEC2Cost = $ExcelWorksheet.columns.Item(13).Rows.Item(100).Text
$cDrive = $ExcelWorksheet.columns.Item(14).Rows.Item(100).Text
$dDrive = $ExcelWorksheet.columns.Item(15).Rows.Item(100).Text
$eDrive = $ExcelWorksheet.columns.Item(16).Rows.Item(100).Text
$data1Volm = $ExcelWorksheet.columns.Item(17).Rows.Item(100).Text
$log1volm = $ExcelWorksheet.columns.Item(18).Rows.Item(100).Text
$data2Volm = $ExcelWorksheet.columns.Item(19).Rows.Item(100).Text
$log2volm = $ExcelWorksheet.columns.Item(20).Rows.Item(100).Text
$data3Volm = $ExcelWorksheet.columns.Item(21).Rows.Item(100).Text
$log3volm = $ExcelWorksheet.columns.Item(22).Rows.Item(100).Text
$backupVolm = $ExcelWorksheet.columns.Item(23).Rows.Item(100).Text
$otherStorage = $ExcelWorksheet.columns.Item(24).Rows.Item(100).Text
$serviceAccountNeeded = $ExcelWorksheet.columns.Item(25).Rows.Item(100).Text
$serviceAccount = $ExcelWorksheet.columns.Item(26).Rows.Item(100).Text
$privateDNSPrefix = $ExcelWorksheet.columns.Item(27).Rows.Item(100).Text
$privateHostedZone = $ExcelWorksheet.columns.Item(28).Rows.Item(100).Text
$certRequired = $ExcelWorksheet.columns.Item(29).Rows.Item(100).Text
$nLoadBalanceerOrApploadbalancer = $ExcelWorksheet.columns.Item(30).Rows.Item(100).Text
$sqlEdition= $ExcelWorksheet.columns.Item(31).Rows.Item(100).Text
$comments = $ExcelWorksheet.columns.Item(32).Rows.Item(100).Text
$hostType = $ExcelWorksheet.columns.Item(33).Rows.Item(100).Text
$dnsHostName = $ExcelWorksheet.columns.Item(34).Rows.Item(100).Text
$dnsSuffix = $ExcelWorksheet.columns.Item(35).Rows.Item(100).Text
$state = $ExcelWorksheet.columns.Item(36).Rows.Item(100).Text
$vueInstall = $ExcelWorksheet.columns.Item(37).Rows.Item(100).Text
Write-Output " ---------------------------Value from Excel Variable added Suceesfully--------------------------"
