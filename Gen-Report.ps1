Function MakeWordReport($index,$hash){
    $template = "C:\Users\tempuser\Documents\TemplateFinding.docx"
    $wd = New-Object –comobject Word.Application
    $doc=$wd.documents.Add($template)
    $newfile="C:\Users\tempuser\Documents\file_$index.docx"
    foreach($key in $hash.keys){
        $objrange = $doc.Bookmarks.Item($key).Range 
        $objrange.Text = $hash[$key]
    }
    $doc.SaveAs([ref]$newfile)
    $doc.Close()
    $wd.Quit()
}

function WorkOnSeverity($RawSeverity){
    $Severity = $RawSeverity.trim()
    $hash["Severity"] = $Severity
}
    
function WorkOnRecommendation($RawRecommendation){
    if($RawRecommendation -like "*Reference:*"){
        $temp =  $RawRecommendation -split "Reference:",2
        $Recommendation = $temp[0].trim()
        $Reference = $temp[1].Trim()
    }
    else{
        $Recommendation = $RawRecommendation.Trim()
        $Reference = ""
    }

    $hash["Recommendation"] = $Recommendation
    $hash["Reference"] = $Reference
}

function WorkOnAffectedResources($RawAffectedResources){
    $AffectedResources = $RawAffectedResources.Trim()
    $hash["AffectedResources"] = $AffectedResources
}

function WorkOnImplication($RawImplication){
    $Implication = $RawImplication.Trim()
    $hash["Implication"] = $Implication

}

function WorkOnObservation($RawObservation){
    $Observation = $RawObservation.trim()
    $hash["Observation"] = $Observation
}

function WorkOnHeading($RawHeading){
    $Heading = $RawHeading.trim()
    $hash["Heading"] = $Heading
}

$objExcel = New-Object -ComObject Excel.Application
$WorkBook = $objExcel.Workbooks.Open("C:\Users\tempuser\Documents\VulnList.xlsx")
$SheetNames = $WorkBook.sheets | Select-Object -Property Name
$WorkSheet = $WorkBook.sheets.item("Network Penetration Testing")
$WorksheetRange = $workSheet.UsedRange
$RowCount = $WorksheetRange.Rows.Count
$ColumnCount = $WorksheetRange.Columns.Count
Write-Host "RowCount:" $RowCount
Write-Host "ColumnCount" $ColumnCount
$ColHeading = 0
$ColObservation = 0
$ColImplication = 0
$ColRecommendation = 0
$ColAffectedResources = 0
$ColSeverity = 0
$hash = @{} # I'll tell u why this is needed in a sec


for($i=1;$i -le $ColumnCount;$i+=1){
    $ColHead = $WorkSheet.cells.Item(1, $i).text
    if($ColHead -like "Heading"){
        Write-Host "[+]Heading Column Found"
        $ColHeading = $i
    }
    if($ColHead -like "Observation"){
        Write-Host "[+]Observation Column Found"
        $ColObservation = $i
    }
    if($ColHead -like "Implication"){
        Write-Host "[+]Implication Column Found"
        $ColImplication = $i
    }
    if($ColHead -like "Recommendation"){
        Write-Host "[+]Recommendation Column Found"
        $ColRecommendation = $i
    }
    if($ColHead -like "Affected Resources"){
        Write-Host "[+]Affected Resources Column Found"
        $ColAffectedResources = $i
    }
    if($ColHead -like "Severity"){
        Write-Host "[+]Severity Column Found"
        $ColSeverity = $i
    }

}

Write-Host "`r`n"
Write-Host "Printing Column Status"
Write-Host "Heading:" $ColHeading
Write-Host "Observation:" $ColObservation
Write-Host "Implication:" $ColImplication
Write-Host "Recommendation:" $ColRecommendation
Write-Host "Affected Resources:" $ColAffectedResources
Write-Host "Severity:" $ColSeverity


for($i=2;$i -le $RowCount; $i+=1){
    $TextHeading = $WorkSheet.cells.Item($i, $ColHeading).text
    $TextObservation = $WorkSheet.cells.Item($i, $ColObservation).text 
    $TextImplication = $WorkSheet.cells.Item($i, $ColImplication).text
    $TextRecommendation = $WorkSheet.cells.Item($i, $ColRecommendation).text
    $TextAffectedResources = $WorkSheet.cells.Item($i, $ColAffectedResources).text
    $TextSeverity = $WorkSheet.cells.Item($i, $ColSeverity).text
    WorkOnHeading $TextHeading
    WorkOnObservation $TextObservation
    WorkOnImplication $TextImplication
    WorkOnRecommendation $TextRecommendation
    WorkOnAffectedResources $TextAffectedResources
    WorkOnSeverity $TextSeverity
    $hash # Just Printing HashTable 
    MakeWordReport $i $hash 
    $hash = @{}
}
