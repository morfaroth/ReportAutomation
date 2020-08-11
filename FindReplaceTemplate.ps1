$objWord = New-Object -comobject Word.Application  
$objWord.Visible = $false

$objDoc = $objWord.Documents.Open("C:\Users\Matthew\Desktop\Template Find and Replace\test.docx") 
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

Import-Csv .\PentestReport-VARIABLES.csv | ForEach-Object {
    Write-Host "$($_.VariableName)"
    $a = $objSelection.Find.Execute($($_.VariableName),$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$($_.Value),$wdReplaceAll) 
}

$objDoc.Save()
$objWord.Quit()