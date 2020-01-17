$Directory_path = Read-Host -Prompt "Введите путь к папке в которой находятся результаты тестирования "
$XMLFiles = Get-ChildItem $Directory_path -Filter *.blist



$Excel = New-Object -comobject Excel.Application
$Excel.Visible = $True
$Excel.DisplayAlerts = $True
$Excel.ScreenUpdating = $True
$Excel.Visible = $True
$CurrentLocation = [string](Get-Location)
$ExcelFilePath = $CurrentLocation + "\Black_list.xlsx"
$WorkBook = $Excel.workbooks.Open($ExcelFilePath)
$WorkSheetName = "Kaspersky Anti-Spam BLACKLIST"
$WorkSheet = $WorkBook.Worksheets.Item($WorkSheetName)
$Cells=$WorkSheet.Cells

$i=2

foreach($XMLfile in $XMLFiles) {
    $ServerName = $XMLfile.BaseName
    $File_full_path = $XMLfile.FullName

    [xml]$XmlDocument = Get-Content -Encoding UTF8 $File_full_path
    $BlackList_Items = $XmlDocument.AntispamBlacklistSettings2.Items.AntispamBlacklistItem

    foreach ($item in $BlackList_Items) {
        Write-Host "++++++++++++++++++++++++++++"
        
        $comment = $item.Comment
        $ID = $item.Id
        $type = $item.ItemType
        $value = $item.ItemValue
        $TimeStamp = $item.ModificationDateTimeUtc
        $ModifiedBy = $item.ModifiedByUser

        # Конвертируем кодировку текста полученного из XML файла для нормального отображения в Excel файле
        #$ModifiedBy = "$ModifiedBy" | ConvertTo-Encoding "UTF-8" "windows-1251"
        #$comment = "$comment" | ConvertTo-Encoding "UTF-8" "windows-1251"

        #------------------------------------------------------------------------------------------------------------
        try {
            $template = 'yyyy-MM-ddTHH:mm:ss.fffffffZ'
            $TimeStamp = [DateTime]::ParseExact($TimeStamp, $template, $null)
        }
        catch {
            try {
                $TimeStamp_NP = $True
                $template = 'yyyy-MM-ddTHH:mm:ss.ffffffZ'
                $TimeStamp = [DateTime]::ParseExact($TimeStamp, $template, $null)
            }
            catch {
                try {
                    $template = 'yyyy-MM-ddTHH:mm:ss.ffffZ'
                    $TimeStamp = [DateTime]::ParseExact($TimeStamp, $template, $null)
                }
                catch {
                    $template = "Timestamp parsing error."
                }
            }
        }
        #------------------------------------------------------------------------------------------------------------

        $Cells.item($i,1) = "$value"
        $Cells.item($i,2) = "$ServerName"
        $Cells.item($i,3) = "$type"
        $Cells.item($i,4) = "$ID"
        $Cells.item($i,5) = "$TimeStamp"
        $Cells.item($i,6) = "$ModifiedBy"
        $Cells.item($i,7) = "$comment"

    $i++
    }

}