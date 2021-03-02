Param(
    $outPath,
    $sourcePath
)

get-process excel | stop-process -force

$sourceFiles = Get-ChildItem -Path $sourcePath #this should be a parameter

$InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

# BLOCK 1: Create and open runspace pool, setup runspaces array with min and max threads
$pool = [RunspaceFactory]::CreateRunspacePool(1, 6, $InitialSessionState, $host)
$pool.ApartmentState = "MTA"
$pool.Open()

$runspaces = $results = @()

# BLOCK 2: Create reusable scriptblock. This is the workhorse of the runspace. Think of it as a function.

$scriptBlock = { #iterate over every workbook in list of workbooks
    Param(
        $file,
        $pageNumber,
        $outpath
    )
    function newMutex {
        Param(
            [string]$mutexId
        )
        $mtx = New-Object System.Threading.Mutex($false, "Global\$mutexId")

        return $mtx

    }

    function waitMutex {
        Param(
            [ref]$mutexRef
        )
        [void]$mutexRef.WaitOne()
    
    }

    function releaseMutex {
        Param(
            [ref]$mutexRef
        )
        [void]$mutexRef.ReleaseMutex()
    }

    $columnIntToAlphaCharHash = @{
        1  = "A"
        2  = "B"
        3  = "C"
        4  = "D"
        5  = "E"
        6  = "F"
        7  = "G"
        8  = "H"
        9  = "I"
        10 = "J"
        11 = "K"
        12 = "L"
        13 = "M"
        14 = "N"
        15 = "O"
        16 = "P"
        17 = "Q"
        18 = "R"
        19 = "S"
        20 = "T"
        21 = "U"
        22 = "V"
        23 = "W"
        24 = "X"
        25 = "Y"
        26 = "Z"
    }

    $excelRoot = New-Object -ComObject Excel.Application
    $workbook = $excelRoot.Workbooks.Open($file.fullname)


    $worksheet = $workbook.sheets.item("Report") #this is bad practice, hard coding the sheet name, but in this case its a known quantity. The sheets property can also return an array of sheets, in this case it doesn't, but you'd want to handle that

    $usedRange = $worksheet.usedRange #gets all known cells to have ever had values from excel

    $colCount = $usedRange.columns.count #gets used column count

    $rowCount = $usedRange.rows.count #gets used row count

    Write-Host "Records in workbook: $rowCount on thread: $pageNumber"
    $dataHashList = New-Object -TypeName 'System.Collections.Generic.List[hashtable]'

    For ($x = 2; $x -le $rowCount; $x++) {
        #iterate over every row in sheet, start at row 2 to exclude headers
        $dataHashTable = @{} #empty k/v store
        For ($i = 1; $i -le $colCount; $i++) {                
            $columnAlpha = $columnIntToAlphaCharHash.$i #get current column letter
            $rangeID = $columnAlpha + $x.toString() #build the range identifier EX: C1 or E5
            $cellText = $worksheet.Range($rangeID).Text #pull text from cell
            $cellText = $cellText.Replace(",", "") #remove any commas from text values and dollar values to allow for correct parsing
            $dataHashTable.Add($columnAlpha, $cellText) #add column name/text data to row k/v store                
        } #for each row, iterate over every column
        $dataHashList.Add($dataHashTable) #add row data table to array of rows
        if ($x % 300 -eq 0) {
            Write-Host "$x records loaded on thread $pageNumber."              
        }
    }


    Write-Host "Thread $pageNumber fully loaded"
       

    $dataList = New-Object -TypeName 'System.Collections.Generic.List[PSCustomObject]'

    #$fancyDate = $true notes unclear, set to true for formatted iso date code
    $4charPostingYear = $true #notes unclear, set to true for 4 character years in postingDate column

    :outerHashLoop
    ForEach ($dataHash in $dataHashList) {
        $outputLine = New-Object -TypeName PSCustomObject
        if ([string]::IsNullOrEmpty($dataHash["F"])) {
            continue outerHashLoop #if date is null, go to next record
        }
        ForEach ($key in $dataHash.keys) {
            Switch ($key) {
                "A" {
                    #build tranID and subsidiary                
              
                    $tranID = $dataHash["G"].SubString(2, 4) #date value
                    $srcPostingDate = $dataHash["F"]
                    $splitSrc = $srcPostingDate.split("/")
                    $tranIDDate = $splitSrc[1].toString() + $splitSrc[0].toString()
                             
                    if ($dataHash["A"] -eq "WA") {
                        $companyNum = "2"
                        $companyIdentifier = "WAHTB"
                    }
                    elseif ($dataHash["A"] -eq "CB") {
                        $companyNum = "3"
                        $companyIdentifier = "CBHTB"
                    }
                    else {
                        continue outerHashLoop #if company value invalid, go to next hashtable record                                        
                    }
                    $tranID = $companyIdentifier + $tranIDDate #add date value to id
                    $outputLine | Add-Member -NotePropertyName "tranId" -NotePropertyValue $tranID
               
                    $outputLine | Add-Member -NotePropertyName "companyIdentifier" -NotePropertyValue $companyIdentifier.substring(0, 2)

                    $outputLine | Add-Member -NotePropertyName "subsidiary" -NotePropertyValue $companyNum
                }
                "B" {
                    $outputLine | Add-Member -NotePropertyName "officeCode" -NotePropertyValue $dataHash["B"]
                }
                "C" {
                    $outputLine | Add-Member -NotePropertyName "journalItemLine_location" -NotePropertyValue $dataHash["C"]
                }
                "D" {
                    # build journalItemLine_account
                    $accountField = $dataHash["D"]
                    $outputLine | Add-Member -NotePropertyName "journalItemLine_account" -NotePropertyValue $accountField
                }
                "F" {
                    #build postingperiod
                    $srcPostingDate = $dataHash["F"]
                    $splitSrc = $srcPostingDate.split("/")
                    if ($4charPostingYear) {
                        $postingDate = Get-Date -format "MMM yyyy" -Year "20$($splitSrc[1])" -Month $splitSrc[0] 
                    }
                    else {
                        $postingDate = Get-Date -format "MMM yy" -Year "20$($splitSrc[1])" -Month $splitSrc[0] 
                    }
                    $orderDateCode = Get-Date -format "yyyyMM" -year "20$($splitSrc[1])" -Month $splitSrc[0]
                    $outputLine | Add-Member -NotePropertyName "orderDateCode" -NotePropertyValue $orderDateCode
                    $outputLine | Add-Member -NotePropertyName "postingperiod" -NotePropertyValue $postingDate.toUpper()
                
                    $lastDayOfMonth = [DateTime]::DaysInMonth("20$($splitSrc[1])", $splitSrc[0])
                
                    $tranDateFormatted = Get-Date -Year "20$($splitSrc[1])" -Month $splitSrc[0] -Day $lastDayOfMonth -format "MM/dd/yyyy"
                    $outputLine | Add-Member -NotePropertyName "tranDate" -NotePropertyValue $tranDateFormatted
                }
                "G" {
                    # bulid trandate - trandate should be iso last day of month
                    #$splitTrandate = $dataHash["G"].Split('.')
                    # $tranDate = $splitTrandate[0]
                
                    # $lastDayOfMonth = [DateTime]::DaysInMonth($tranDate.subString(0, 4), $trandate.substring(5, 2))
                
                    # $tranDateFormatted = Get-Date -Year $tranDate.substring(0, 4) -Month $tranDate.substring(5, 2) -Day $lastDayOfMonth -format "MM/dd/yyyy"
                    # $outputLine | Add-Member -NotePropertyName "tranDate" -NotePropertyValue $tranDateFormatted
                    $outputLine | Add-Member -NotePropertyName "datedoc" -NotePropertyValue $dataHash["G"]
                }
                "H" {                    
                    #build memo field
                    $dirtyDescription = $dataHash["H"]
                    $cleanDescription = $dirtyDescription -replace '\s+',' ' #regex to remove unnecessary whitespace
                    $outputLine | Add-Member -NotePropertyName "memo" -NotePropertyValue $cleanDescription
                }
                "J" {
                    #build journalItemLine_debitAmount
                    $outputLine | Add-Member -NotePropertyName "journalItemLine_debitAmount" -NotePropertyValue $dataHash["J"]
                }
                "K" {
                    #build journalItemLine_creditAmount
                    $outputLine | Add-Member -NotePropertyName "journalItemLine_creditAmount" -NotePropertyValue $dataHash["K"]
                }
                "M"{
                    $outputLine | Add-Member -NotePropertyName "lineNetActivity" -NotePropertyValue $dataHash["M"]
                }
                "N"{
                    $outputLine | Add-Member -NotePropertyName "lineEndingBalance" -NotePropertyValue $dataHash["N"]
                }
                "O"{
                    $outputLine | Add-Member -NotePropertyName "lineMoProductName" -NotePropertyValue $dataHash["O"]
                }
                "P"{
                    $outputLine | Add-Member -NotePropertyName "lineMoProductCode" -NotePropertyValue $dataHash["P"]
                }
                "R"{
                    $outputLine | Add-Member -NotePropertyName "lineName" -NotePropertyValue $dataHash["R"]
                }
                "S"{
                    $outputLine | Add-Member -NotePropertyName "lineMoSource" -NotePropertyValue $dataHash["S"]
                }
                "T" {
                    #build MO_ID
                    $outputLine | Add-Member -NotePropertyName "MO_ID" -NotePropertyValue $dataHash["T"]
                }
                "U" {
                    #build MO_ID
                    $outputLine | Add-Member -NotePropertyName "tbd" -NotePropertyValue $dataHash["U"]
                }

            }
        }

        #add non-source values 
        # $outputLine | Add-Member -NotePropertyName "journalItemLine_class" -NotePropertyValue ""
        # $outputLine | Add-Member -NotePropertyName "journalItemLine_department" -NotePropertyValue ""
        # $outputLine | Add-Member -NotePropertyName "isdferred" -NotePropertyValue "FALSE"

        $dataList.Add($outputLine)
    }

    Write-Host "Generating files for thread $pageNumber"
    ForEach ($item in $dataList) {
    
        $mutex = newMutex -mutexID "$($item.companyIdentifier)-$($item.orderDateCode)"

        $fileOutPath = "$outpath\FileName $($item.companyIdentifier) $($item.orderDateCode).csv"
        if (!(test-path -path $fileOutPath)) {
            #year/month file doesn't exist yet, creating and adding headers
            waitMutex -mutexRef $mutex

            New-Item -ItemType File -Path $outpath -Name "File Name $($item.companyIdentifier) $($item.orderDateCode).csv" | out-null
            Add-Content -path $fileOutPath -Value "tranId,,subsidiary,trandate,postingperiod,,journalItemLine_location,journalItemLine_account,,memo,journalItemLine_debitAmount,journalItemLine_creditAmount,,,,,,,MO_ID,,"
            Add-Content -path $fileOutPath -Value "External ID,Subsidiary,Sub Int ID,Tran Date,Posting Period,Line Office Ext ID,Line Office,Line Account Ext ID,Line MO Date.Doc,Line Memo,Line Debit,Line Credit,Line Net Activity,Line Ending Balance,Line MO Product Name,Line MO Product Code,Line Name,Line MO Source,Line MO Reference,TBD"
            Add-Content -path $fileOutPath -Value ",,,,,,,,,Free-Form Text,,,Currency,Currency,Type TBD,Type TBD,,Free-Form Text,Free-Form Text,,"

            releaseMutex -mutexRef $mutex
        }
        
            waitMutex -mutexRef $mutex
            Add-Content -Path $fileOutPath -Value "$($item.tranId),$($item.companyIdentifier.SubString(0,2)),$($item.subsidiary),$($item.trandate),$($item.postingperiod),$($item.officeCode),$($item.journalItemLine_location),$($item.journalItemLine_account),$($item.datedoc),$($item.memo),$($item.journalItemLine_debitAmount),$($item.journalItemLine_creditAmount),$($item.lineNetActivity),$($item.lineEndingBalance),$($item.lineMoProductName),$($item.lineMoProductCode),$($item.lineName),$($item.lineMoSource),$($item.MO_ID),$($line.tbd)"
            releaseMutex -mutexRef $mutex
        
    }

    $excelRoot.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelRoot)
    Remove-Variable excelRoot
}


# BLOCK 3: Create runspace and add to runspace pool
$i = 1
ForEach ($file in $sourceFiles) {

    $runspace = [PowerShell]::Create()
    [void]$runspace.AddScript($scriptblock)
    [void]$runspace.AddArgument($file) #file
    [void]$runspace.AddArgument($i) #$pageNumber
    [void]$runspace.AddArgument($outPath)
   
    $runspace.RunspacePool = $pool

    # BLOCK 4: Add runspace to runspaces collection and "start" it
    # Asynchronously runs the commands of the PowerShell object pipeline
    $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
    $i = $i + 1
}

# BLOCK 5: Wait for runspaces to finish
while ($runspaces.Status.IsCompleted -notcontains $true) {}


# BLOCK 6: Clean up
foreach ($runspace in $runspaces ) {
    # EndInvoke method retrieves the results of the asynchronous call
    $results += $runspace.Pipe.EndInvoke($runspace.Status)
    $runspace.Pipe.Dispose()
}
    
$pool.Close() 
$pool.Dispose()

Write-Host ($results | out-string)


