#Snippet...

$wdTypes = Add-Type -AssemblyName "Microsoft.Office.Interop.Word" -PassThru
$wdSaveFormat = $wdTypes | WHere Name -eq "wdSaveFormat"
# Get all names: [enum]::getnames($wdSaveFormat)
# results in wdFormatPDF
$wdSaveFormat::wdFormatPDF.Value__ 
