$Directory_path = Read-Host -Prompt "Enter the path to the folder containing the files with the extension * .blist: "
$XMLFiles = Get-ChildItem $Directory_path -Filter *.blist 

if ($XMLFiles.Count -gt 1) {

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

            #------------------------------------------------------------------------------------------------------------
            $difference = (($TimeStamp.LastIndexOf("Z")) - ($TimeStamp.LastIndexOf(".")) - 1)
            $template = ''

            if ((($TimeStamp.ToCharArray()) -contains [char]'Z') -and (($TimeStamp.ToCharArray()) -contains [char]'.'))  {
    
            if (($difference -gt 0) -and ($difference -le 7)) {  
        
                switch ($difference) {
                    7 {$template = 'yyyy-MM-ddTHH:mm:ss.fffffffZ'}
                    6 {$template = 'yyyy-MM-ddTHH:mm:ss.ffffffZ'}
                    5 {$template = 'yyyy-MM-ddTHH:mm:ss.fffffZ'}
                    4 {$template = 'yyyy-MM-ddTHH:mm:ss.ffffZ'}
                    3 {$template = 'yyyy-MM-ddTHH:mm:ss.fffZ'}
                    2 {$template = 'yyyy-MM-ddTHH:mm:ss.ffZ'}
                    1 {$template = 'yyyy-MM-ddTHH:mm:ss.fZ'}
                }

                try {
                    $TimeStamp = [DateTime]::ParseExact($TimeStamp, $template, $null)
                }
                catch {
                    $TimeStamp = "Error parsing timestamp. Timestamp = $TimeStamp"
                }
            }

            } else {
                $TimeStamp = "Error parsing timestamp. Timestamp = $TimeStamp" 
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
} else {
    Write-Host "ERROR: The specified directory does not contain files matching the * .blck mask!" -ForegroundColor Red
}