Add-Type -AssemblyName System.Windows.Forms
$objWord = New-Object -comobject Word.Application  
$objWord.Visible = $false

$FindingsFile = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }

$null = $FindingsFile.ShowDialog()

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
$retOutput = @()

Import-Csv $FindingsFile.FileName | ForEach-Object {
    $objDoc = $objWord.Documents.Open($PSScriptRoot+"\FindingTemplate.docx") 
    $objSelection = $objWord.Selection
   

    $objrange = $objDoc.Bookmarks.Item("Name").Range 
	$objrange.Text = $($_.Name)

	$objrange = $objDoc.Bookmarks.Item("Risk").Range 
	$objrange.Text = $($_.Risk)
	
	$objrange = $objDoc.Bookmarks.Item("Description").Range 
	$objrange.Text = $($_.Description)
    
	$objrange = $objDoc.Bookmarks.Item("Impact").Range 
	$objrange.Text = $($_.Impact)

	$objrange = $objDoc.Bookmarks.Item("AffectedResources").Range 
	$objrange.Text = $($_.AffectedResources)

	$objrange = $objDoc.Bookmarks.Item("Recommendation").Range 
	$objrange.Text = $($_.Recommendation)

	$objrange = $objDoc.Bookmarks.Item("Scenario").Range 
	$objrange.Text = $($_.Scenario)

    $name = $PSScriptRoot+"\$($_.Name).docx"
    $findingDocuments += "$($name)"
    $objDoc.Saveas([ref]$name,[ref]$SaveFormat::wdFormatDocument)
    $objDoc.Close()
}

$objDoc = $objWord.Documents.Add()
$objSelection = $objWord.Selection
foreach($finding in $findingDocuments){
    $objSelection.TypeParagraph()
    $objSelection.InsertFile($finding)
	Remove-Item -Path $finding
}
$name = $PSScriptRoot+"\allfindings.docx"
$objDoc.Saveas([ref]$name,[ref]$SaveFormat::wdFormatDocument)
$objWord.Quit()