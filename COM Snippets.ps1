#region Credentials
Try {
	$tmp = Import-Clixml $Env:USERPROFILE\Credential.xml -ErrorAction Stop
	$Cred = New-Object System.Management.Automation.PSCredential $tmp['UserName'],($tmp['Password'] | ConvertTo-SecureString)
} 
Catch {
	$Cred = Get-Credential "$ENV:USERDNSDOMAIN\Admin Credential"
    # Save the credentials for future use....
	@{'UserName'=$Cred.UserName;'Password'=($Cred.Password | ConvertFrom-SecureString)} | 
        Export-Clixml "$Env:USERPROFILE\Credential.xml"
}
<#
 !!!!! WARNING !!!!
 $Credential.GetNetworkCredential().Password will decrypt the password on the machine where it is stored.
#>
#endregion Credentials

#region Outlook
[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Outlook") | out-null
$olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
$OlClass = "Microsoft.Office.Interop.Outlook.OlObjectClass" -as [type]
$OlSaveAs = "Microsoft.Office.Interop.Outlook.OlSaveAsType" -as [type]
$OlBodyFormat = "Microsoft.Office.Interop.Outlook.OlBodyFormat" -as [type]

$Oapp = new-object -comobject outlook.application
$namespace = $Oapp.GetNameSpace("MAPI")
$namespace.Folders | Select -ExpandProperty Name
$Outlook = $namespace.Folders["art.beane@outlook.com"]
$Outlook.Folders | Format-Table Name,UnreadItemCount,@{N='Items';E={$_.Items.Count}} -AutoSize
$Folder = $Outlook.Folders.Item("Inbox")

$Folder.Items.item(1) | Get-Member -MemberType *Property | Sort Name
$Folder.Items | Group SenderName | Where Count -GT 5
$msgs = $Folder.Items | Where {$_.SenderEmailAddress -match "brewery@saintarnoldbrewingcompany"}
$msgs | Format-Table SentOn,SenderName,Subject -AutoSize
$Folder.Items.Item(1) | Format-Table SentOn,SenderName,Subject -AutoSize
$Folder.Folders
$Outlook.Folders | Format-Table Name,UnreadItemCount,@{N='Items';E={$_.Items.Count}}

#endregion Outlook

#region Visio
$visTypes = Add-Type -AssemblyName ‘Microsoft.Office.Interop.Visio, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c' -PassThru
$visColors = $visTypes | Where Name -eq 'VisDefaultColors'
$visCells = $visTypes | Where Name -eq 'visCellIndices'
$visPaper = $visTypes | Where Name -eq 'VisPaperSizes'
$visIndices = $visTypes | where Name -eq 'visSectionIndices'
$visRows = $visTypes | Where {Name -eq 'visRowIndices'
#    [enum]::GetNames($visColors)
$Visio = New-Object -ComObject Visio.Application
$docs = $Visio.Documents
$Doc = $docs.Add("Basic Network Diagram (US Units).vst")
$Doc.PaperSize = $visPaper::visPaperSizeE
# Set the active page of the document to page 1
$pages = $Visio.ActiveDocument.Pages
# Load a set of stencils and select one to drop
$shpFile = "Virtualization Visio Stencil.vss"
$stnPath = Join-Path (Join-Path ([system.Environment]::GetFolderPath('MyDocuments')) "My Shapes") $shpFile
$stencil = $Visio.Documents.Add($stnPath)
$stnVCenter = $stencil.Masters.Item("Virtual Center Management Console")
$HostShp = $stencil.Masters.Item("ESX Host")
$MSShp = $stencil.Masters.Item("Microsoft Server")
$LXSxp = $stencil.Masters.Item("Linux Server")
$OtherShp =  $stencil.Masters.Item("Other Server")
$CluShp = $stencil.Masters.Item("Cluster")
$DCShp = $stencil.Masters.Item("Datacenter")

function Add-VisioObject {
    Param ($mastObj, $item, [string]$Power="None")
 		Write-Host "Adding $item ($x, $y), $Power"
  		$shpObj = $page.Drop($mastObj, $x, $y)
  		$shpObj.Text = $item
        Switch -Regex ($Power) {
            "None" {$Color =  $visColors::visBlack; Break}
            "ON" {$Color = $visColors::visDarkGreen; break}
            "OFF" {$Color = $visColors::visDarkRed; break}
            Default {$Color = $visColors::visBlack}
        }
        $shpObj.Characters.CharProps($visCells::visCharacterColor) = $Color
        $shpObj.CellsSRC($VisIndices::visSectionCharacter,$visRows::visRowCharacter,$visCells::visCharacterFont).ResultIU = 62 # Cambria
        $shpObj.Characters.CharProps($visCells::visCharacterSize)=10
		return $shpObj
 }

function Connect-VisioObject {
	Param  ($firstObj, $secondObj)
    $shpConn = $page.Drop($page.Application.ConnectorToolDataObject, 0, 0)
	#// Connect its Begin to the 'From' shape:
	$connectBegin = $shpConn.CellsU("BeginX").GlueTo($firstObj.CellsU("PinX"))
	#// Connect its End to the 'To' shape:
	$connectEnd = $shpConn.CellsU("EndX").GlueTo($secondObj.CellsU("PinX"))
}

Function Draw-TextBox {
    Param ($Shape,$Text)
    $x = $Shape.CellsU('PinX').ResultIU - 1.0
    $y = $Shape.CellsU('PinY').ResultIU - 1.25
    $tb = $Page.DrawRectangle($x,$y,$x+2.0,$y+0.5)
    $tb.LineStyle = "Text Only"
    $tb.FillStyle = "Text Only"
    $tb.Text = $Text
    $tb.CellsSRC($VisIndices::visSectionObject,$visRows::visRowFill,$visCells::visFillBkgnd).FormulaU='THEMEGUARD(SHADE(FillForegnd,LUMDIFF(THEME("FillColor"),THEME("FillColor2"))))'
    $tb.CellsSRC($VisIndices::visSectionObject,$visRows::visRowFill,$visCells::visFillForegnd).FormulaU='THEMEGUARD(MSOTINT(THEME("AccentColor4"),80))'
    $tb.CellsSRC($VisIndices::visSectionObject,$visRows::visRowFill,$visCells::visFillPattern).ResultIU=1
    $tb.CellsSRC($VisIndices::visSectionObject,$visRows::visRowLine,$visCells::visLinePattern).ResultIU=1
    $ArialNarrow = 1..($Doc.Fonts.Count) | foreach {If ($Doc.Fonts.Item($_).Name -eq "Arial Narrow") {$Doc.Fonts.Item($_).Index; Break}}
    $tb.CellsSRC($VisIndices::visSectionCharacter,$visRows::visRowCharacter,$visCells::visCharacterFont).ResultIU = $ArialNarrow
    #$tb.Characters.CharProps($visCells::visCharacterSize)=9
    $tb.Characters.CharProps(7)=9
}
#endregion Viso

#region Excel
$XLTypes = Add-Type -AssemblyName "Microsoft.Office.Interop.Excel" -PassThru
$HAlign = $xltypes | Where-Object Name -eq 'XlHAlign'
$VAlign = $xltypes | Where-Object Name -eq 'XlVAlign'
$FileFormat = $xltypes | Where-Object Name -eq 'XlFileFormat'
$SheetType = $xltypes | Where-Object Name -eq 'XlSheetType'
$xlVAlignTop = $VAlign::xlValignTop # -4160
$XlHAlignCenter = $HAlign::xlHAlignCenter #-4108
$XlWorksheet = $SheetType::xlWorkSheet # -4167
$XlWorkbookDefault = $FileFormat::xlWorkbookDefault # 51 (.xlsx)
$XL = New-Object -ComObject Excel.Application
$XL.SheetsInNewWorkbook = 1
$XLWB=$XL.Workbooks.Add()
$XLWB.Activate()
$XL.Visible = $True
Function Add-CSVSheet {
	Param ([String]$Name,$Collection)
    if ($Collection.GetType().Name -eq "HashTable") {
        $Rows = @()
        $Collection | foreach {
            $Row = "" | Select $_.GetEnumerator().Name
            $_.GetEnumerator() | foreach {$Row.($_.Name) = $_.Value}
            $Rows += $Row
        } 
    } ELSE {$Rows = $Collection}
    $Rows | Export-Csv "$ScriptRoot\Temp\$Name.csv" -NoTypeInformation
    $XLWB2=$XL.Workbooks.Open("$ScriptRoot\Temp\$Name.csv")
    $Sh = $XLWB2.Worksheets.Item(1)
    $Sheet = $XLWB.Worksheets.Item(1)
    $sh.Copy($Sheet)
    $XLWB2.Close()
    $Labels = $Rows[0] | Get-Member -MemberType *Property
    Set-XLSheetView $Labels
} # end function Add-CSVSheet
Function Set-XLSheetView {
    Param ($Labels)
    $Rank = [int]([math]::Floor($Labels.Count/26))
    $File = $Labels.Count % 26
    If ($File -gt 0) {$l2 = "$([char](64+$File))"} Else {$l2 = 'Z'; $Rank--}
    If ($Rank -gt 0) {$l1 = "$([char](64+$Rank))"} Else {$l1 = ''}
    $limit = "$l1$l2" + "1"
    $Window = $XL.ActiveWindow
    $Window.Zoom = 80
    $Window.SplitRow = 1
    $Window.SplitColumn = 1
    $Window.FreezePanes = $true
    $Sheet = $XLWB.ActiveSheet
    $Sheet.Range("A1:$Limit").EntireColumn.ColumnWidth = 75 # this effectively sets the maximum column width
    $Sheet.Range("A1:$Limit").EntireColumn.WrapText = $TRUE
    $objRange = $Sheet.UsedRange
    $objRange.VerticalAlignment = $xlVAlignTop
    $Ignore=$objRange.Rows.Autofit()
    $Ignore=$objRange.EntireColumn.Autofit()
    $Range = $Sheet.Range("A1:$Limit")
	$Range.Font.Name = "Cambria"
	$Range.Font.Bold = $True
	$Range.HorizontalAlignment = $XlHAlignCenter
} #End Set-XLSheetView

#endregion Excel
