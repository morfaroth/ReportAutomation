$objWord = New-Object -comobject Word.Application  
$objWord.Visible = $false

 
 
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
$findingDocuments = @()

Import-Csv .\findings.csv | ForEach-Object {
    $objDoc = $objWord.Documents.Open("C:\Users\Matthew\Desktop\Template Find and Replace\FindingTemplate.docx") 
    $objSelection = $objWord.Selection

    $a = $objSelection.Find.Execute("<Name>",$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$($_.Name),$wdReplaceAll) 
    
    $a = $objSelection.Find.Execute("<Risk>",$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$($_.Risk),$wdReplaceAll)
    
    $a = $objSelection.Find.Execute("<Description>",$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$($_.Description),$wdReplaceAll)
    
    $a = $objSelection.Find.Execute("<Impact>",$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$($_.Impact),$wdReplaceAll) 

    $a = $objSelection.Find.Execute("<AffectedResources>",$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$($_.AffectedResources),$wdReplaceAll) 

    $a = $objSelection.Find.Execute("<Recommendation>",$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$($_.Recommendation),$wdReplaceAll) 

    $a = $objSelection.Find.Execute("<Scenario>",$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$($_.Scenario),$wdReplaceAll) 

    $name = "C:\Users\Matthew\Desktop\Template Find and Replace\$($_.Name).docx"
    $findingDocuments += "$($name)"
    $objDoc.Saveas([ref]$name,[ref]$SaveFormat::wdFormatDocument)
}

$objDoc = $objWord.Documents.Add()
$objSelection = $objWord.Selection
foreach($finding in $findingDocuments){
    $objSelection.TypeParagraph()
    $objSelection.InsertFile($finding)
}
$name = "C:\Users\Matthew\Desktop\Template Find and Replace\allfindings.docx"
$objDoc.Saveas([ref]$name,[ref]$SaveFormat::wdFormatDocument)
$objWord.Quit()