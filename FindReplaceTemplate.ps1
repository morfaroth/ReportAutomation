$objWord = New-Object -comobject Word.Application  
$objWord.Visible = $false

$TemplateFile = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
$TemplateFile.filter = "DOCX (*.docx)| *.docx"
$null = $TemplateFile.ShowDialog()


$objDoc = $objWord.Documents.Open($TemplateFile.FileName) 
$objSelection = $objWord.Selection 
 
$MatchCase = $False 
$MatchWholeWord = $true
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $wdFindContinue 
$Format = $False 
$wdReplaceNone = 0 
$wdFindContinue = 1 
$wdReplaceAll = 2

$VariableFile = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
$VariableFile.filter = "CSV (*.csv)| *.csv"
$null = $VariableFile.ShowDialog()

Import-Csv $VariableFile.FileName | ForEach-Object {
    $a = $objSelection.Find.Execute($($_.VariableName),$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$($_.Value),$wdReplaceAll) 
}

$objDoc.Save()
$objWord.Quit()