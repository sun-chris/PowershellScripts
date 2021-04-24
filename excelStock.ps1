#The purpose of this script is to Download a megalist of all active stock tickers from AlphaVantage, load that list into Excel, 
#automate the utilization of excel's built-in stock data connections to get basic metadata about all > 10000 stocks, then save back to CSV to use in other scripts;
#With the intention of using this basic stock data to perform initial filtering on which stocks to request full datasets for

#Windows Roadblock #1, Unless you disable the progress bar, it drags the download time of this ~50kb file up from a few seconds, to 5-10 minutes
$ProgressPreference = 'SilentlyContinue'

#Download megalist of active stock tickers from alphavantage
Write-Host "Downloading megalist of active stock tickers"
Invoke-WebRequest -URI "https://www.alphavantage.co/query?function=LISTING_STATUS&apikey=demo" -OutFile '.\raw_stock.csv'

#Get count of lines in CSV
$k = (Get-Content ".\raw_stock.csv" | Measure-Object -Line).Lines
Write-Host $k "stock in sheet"

#Launch Excel and open downloaded csv
$excel =  New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$datapath = [string](Get-Location) + "\raw_stock.csv"
$book = $excel.Workbooks.Open($datapath)
$sheet = $book.Sheets.Item(1)

#Windows Roadblock #2, Excel has an arbitrary limit on the amount of cells that can be converted to a linked data type at once. To work around this, it must be done in 2 batches
$kHalf = [math]::floor($k/2)
#Convert stock tickers in column 1 to Stock Data objects
$r1 = $sheet.range("A2:A${kHalf}")
$r2 = $sheet.range("A${kHalf}:A${k}")
Write-Host "Converting first batch of tickers to linked data type"
$r1.ConvertToLinkedDataType(268435456,"en-US")
Write-Host "Converting second batch of tickers to linked data type"
$r2.ConvertToLinkedDataType(268435456,"en-US")

#Windows Roadblock #3: Getting the properties of a Linked Data Type are an "Insider Only" thing apparently. So I've hard-coded all the fields here
$fields = ("52 Week High","52 Week Low","Beta","Change","Change % (Extended hours)","Change (%)","Change (Extended Hours)","Currency","Description","Employees","Exchange","Exchange abbreviation","Headquarters","High","Industry","Instrument Type","Last trade time","Low","Market cap","Name","Official name","Open","P/E","Previous close","Price","Price (Extended hours)","Shares outstanding","Ticker symbol","Volume","Volume average","Year incorporated")

#Populate columns of sheet with values from the linked data cell
$cptr = 8
Write-Host "Expanding sheet to have all data, before saving to CSV"
foreach ($f in $fields){
    #Add header to column
    $sheet.Cells(1,$cptr).Value = $f
    #Get range of cells to be set with integer values, then apply the formula to them
    $Sheet.Range($sheet.Cells(2,$cptr),$sheet.Cells($k,$cptr)).Value = "=A2:A$k.[$f]"
    #Increment column pointer
    $cptr += 1
}

#Rename a few columns to avoid dupes
$sheet.Cells(1,1).Value = "xlObject"
$sheet.Cells(1,2).Value = "nameAlpha"
$sheet.Cells(1,3).Value = "exchangeAlpha"

#Lift this data out of excel, and back into CSV format before exiting Excel
$book.SaveAs([string](Get-Location) + "\stockMetadata.csv",6)
$excel.Quit()

#Parse through it again to replace the "#FIELD!" cells with "" (Thing that excel puts in for null values)
Write-Host "Cleaning up Excel data"
$fullcsv = Import-Csv ".\stockMetadata.csv"
foreach ($row in $fullcsv){
    foreach ($f in $fields){
        if ($row.$f -eq "#FIELD!"){
            $row.$f = ""
        } 
    }
}
$fullcsv | Export-Csv -Path ".\stockMetadata.csv"

#Windows Roadblock #4: Remove "#TYPE System.Management.Automation.PSCustomObject" line from the top of exported CSV (???)
Get-Content '.\stockMetadata.csv' | Select-Object -Skip 1 | Out-File ".\stockMetadataFinal.csv"

#Clean up
Remove-Item ".\raw_stock.csv"
Remove-Item ".\stockMetadata.csv"
Write-Host "Done!"
