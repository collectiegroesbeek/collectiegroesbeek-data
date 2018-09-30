# Powershell script that converts Excel files to csv.
# The csv output is Tab delimited UTF-16LE with leading Byte Order Mark.

# Put in the path here:
$path = ''

$excel = new-object -ComObject "Excel.Application"
$excel.DisplayAlerts = $True
$excel.Visible = $false

Get-ChildItem $path -Filter *.xlsx | Foreach-Object{
    'processing ' + $_.FullName
    $wb = $excel.Workbooks.Open($_.FullName)
    $dst_file = $path + "\" + $_.BaseName + ".csv"
    if([System.IO.File]::Exists($dst_file)){
        Remove-Item –path $dst_file
    }
    $wb.SaveAs($dst_file, 42)  # 42 is Tab delimited UTF-16LE with leading Byte Order Mark
    'saved ' + $dst_file
    $wb.Close($True)
    Start-Sleep -Seconds 2
}

$excel.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
