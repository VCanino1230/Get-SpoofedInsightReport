#Connect-ExchangeOnline

$Spoofed_Data_Allow = Get-SpoofIntelligenceInsight | Select-Object spoofeduser, sendinginfrastructure, spooftype, messagecount, lastseen, action | Where-Object {$_.Action -eq "Allow"}| Sort-Object action 
$Spoofed_Data_Block = Get-SpoofIntelligenceInsight | Select-Object spoofeduser, sendinginfrastructure, spooftype, messagecount, lastseen, action | Where-Object {$_.Action -eq "Block"}| Sort-Object action

$Spoofed_Insight_Report = New-Object -ComObject excel.application
$Spoofed_Insight_Report.visible = $true
$Spoofed_Insight_Workbook = $Spoofed_Insight_Report.workbooks.add()
$Spoofed_Insight_Workbook.worksheets.add()
$Spoofed_Insight_Workbook.worksheets.add()


$Spoofed_Insight_Allow_Sheet = $Spoofed_Insight_Workbook.worksheets.item(1)
$Spoofed_Insight_Block_Sheet = $Spoofed_Insight_Workbook.worksheets.item(2)
$Spoofed_Insight_Allow_Sheet.name = 'Allowed Spoofed Domains'
$Spoofed_Insight_Block_Sheet.name = 'Blocked Spoofed Domains'

$Spoofed_Insight_Allow_Sheet.cells.item(1,1) = 'Spoofed Insight Report - Allow'
$Spoofed_Insight_Allow_Sheet.cells.item(2,1) = 'Spoofed User'
$Spoofed_Insight_Allow_Sheet.cells.item(2,2) = 'Sending Infrastructure'
$Spoofed_Insight_Allow_Sheet.cells.item(2,3) = 'Spoof Type'
$Spoofed_Insight_Allow_Sheet.cells.item(2,4) = 'Message Count'
$Spoofed_Insight_Allow_Sheet.cells.item(2,5) = 'Last Seen'
$Spoofed_Insight_Allow_Sheet.cells.item(2,6) = 'Action'

$Spoofed_Insight_Block_Sheet.cells.item(1,1) = 'Spoofed Insight Report - Block'
$Spoofed_Insight_Block_Sheet.cells.item(2,1) = 'Spoofed User'
$Spoofed_Insight_Block_Sheet.cells.item(2,2) = 'Sending Infrastructure'
$Spoofed_Insight_Block_Sheet.cells.item(2,3) = 'Spoof Type'
$Spoofed_Insight_Block_Sheet.cells.item(2,4) = 'Message Count'
$Spoofed_Insight_Block_Sheet.cells.item(2,5) = 'Last Seen'
$Spoofed_Insight_Block_Sheet.cells.item(2,6) = 'Action'


$count = 3
foreach($domain in $Spoofed_Data_Allow){
    $Spoofed_Insight_Allow_Sheet.cells.item($count,1) = $domain.SpoofedUser
    $Spoofed_Insight_Allow_Sheet.cells.item($count,2) = $domain.sendinginfrastructure
    $Spoofed_Insight_Allow_Sheet.cells.item($count,3) = $domain.SpoofType
    $Spoofed_Insight_Allow_Sheet.cells.item($count,4) = $domain.MessageCount
    $Spoofed_Insight_Allow_Sheet.cells.item($count,5) = $domain.LastSeen
    #$Spoofed_Insight_Allow_Sheet.cells.item($count,6) = $domain.Action
    $count += 1
}

$count = 3
foreach($domain in $Spoofed_Data_Block){
    $Spoofed_Insight_Block_Sheet.cells.item($count,1) = $domain.SpoofedUser
    $Spoofed_Insight_Block_Sheet.cells.item($count,2) = $domain.sendinginfrastructure
    $Spoofed_Insight_Block_Sheet.cells.item($count,3) = $domain.SpoofType
    $Spoofed_Insight_Block_Sheet.cells.item($count,4) = $domain.MessageCount
    $Spoofed_Insight_Block_Sheet.cells.item($count,5) = $domain.LastSeen
    #$Spoofed_Insight_Block_Sheet.cells.item($count,6) = $domain.Action
    $count += 1
}


#automatically sizes columns in each sheet
$Spoofed_Insight_Allow_Sheet.columns.AutoFit()
$Spoofed_Insight_Block_Sheet.columns.AutoFit()

#file path specifying saving workbook in .xlsx format
#$Report_FilePath = 'C:\Users\vcanino\OneDrive - Nylok LLC\Documents\Reports\MailboxDelegationReport4.xlsx'
$Report_FilePath = 'C:\Users\vcanino\OneDrive - Nylok LLC\Documents\Reports\Spoofed_Insight_Report.xlsx'

#saves Workbook in the file path specified
$Spoofed_Insight_Report.displayalerts = $false
$Spoofed_Insight_Workbook.saveas($Report_FilePath)
$Spoofed_Insight_Report.displayalerts = $true