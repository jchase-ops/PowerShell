<#
.SYNOPSIS
This function imports a file, sorts the data, and writes the data to an Excel table.

.DESCRIPTION
This function sorts and formats an Excel table.  It can import a CSV to use for the table, or it can read a previous Excel table and recreate it on a new worksheet.

.PARAMETER ExcelPath
The filepath to the Excel document.

.PARAMETER ImportPath
The filepath to the CSV for import.

.PARAMETER Headers
An array of strings to be used as headers for the table.

.EXAMPLE
$ExcelPath = "$HOME\ExcelFile.xlsx"
Format-ExcelTable -ExcelPath $ExcelPath

This will:
- open an existing Excel file
- read the data from the table on the first worksheet
- sort and rewrite the data as a new table on a new worksheet
- then save and close the file

.EXAMPLE
$ExcelPath = "$HOME\ExcelFile.xlsx"
$ImportPath = "$HOME\ImportFile.csv"
Format-ExcelTable -ExcelPath $ExcelPath -ImportPath $ImportPath

This will:
- import data from a CSV
- sort and write the data as a new table on a new worksheet
- then save and close the file

.EXAMPLE
$ExcelPath = "$HOME\ExcelFile.xlsx"
$ImportPath = "$HOME\ImportFile.csv"
$Headers = 'Hostname', 'IP', 'Application', 'Tier'
Format-ExcelTable -ExcelPath $ExcelPath -ImportPath $ImportPath -Headers $Headers

This will:
- import data from a CSV
- sort and write the data as a new table on a new worksheet
- attempt to only use the provided $Headers
- will combine the provided $Headers with all remaining object properties
- then save and close the file

.INPUTS
System.String
System.String[]

.OUTPUTS
None.  This function does not return output to the console.

.NOTES
Author: Joshua Chase
DateModified: 2 Jun 2024
#>

function Format-ExcelTable {

    [CmdletBinding(DefaultParameterSetName = 'Import')]

    Param(

        [Parameter(Mandatory, Position = 0, ParameterSetName = 'NoImport')]
        [Parameter(Position = 0, ParameterSetName = 'Import')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ExcelPath,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'Import')]
        [ValidateScript({ Test-Path -Path $_ })]
        [System.String]
        $ImportPath,

        [Parameter(Position = 2, ParameterSetName = 'Import')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Headers
    )

    $numberToAlphabet = @{
        1  = 'A'
        2  = 'B'
        3  = 'C'
        4  = 'D'
        5  = 'E'
        6  = 'F'
        7  = 'G'
        8  = 'H'
        9  = 'I'
        10 = 'J'
        11 = 'K'
        12 = 'L'
        13 = 'M'
        14 = 'N'
        15 = 'O'
        16 = 'P'
        17 = 'Q'
        18 = 'R'
        19 = 'S'
        20 = 'T'
        21 = 'U'
        22 = 'V'
        23 = 'W'
        24 = 'X'
        25 = 'Y'
        26 = 'Z'
    }

    switch ($PSCmdlet.ParameterSetName) {
        'Import' {

            if (!($Headers)) {
                $dataFromImport = Import-Csv -Path $ImportPath
                $propertyList = [System.Collections.Generic.List[System.String]]::New()
                $propertyList.Add('Hostname')
                $dataFromImport | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -ne 'Hostname' } | Select-Object -ExpandProperty Name | Sort-Object -Unique | ForEach-Object { $propertyList.Add($_) }
                $sortedData = $dataFromImport | Select-Object -Property $propertyList | Sort-Object -Property Hostname
            }
            else {
                $dataFromImport = Import-Csv -Path $ImportPath -Header $Headers
                if ($Headers[0] -ne 'Hostname') {
                    $propertyList = [System.Collections.Generic.List[System.String]]::New()
                    $propertyList.Add('Hostname')
                    $tempHeaderList = $Headers | Where-Object { $_ -ne 'Hostname' } | Sort-Object
                    $tempPropertyList = $dataFromImport | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -ne 'Hostname' } | Select-Object -ExpandProperty Name | Sort-Object
                    Compare-Object -ReferenceObject $tempPropertyList -DifferenceObject $tempHeaderList -IncludeEqual | Where-Object { $_.SideIndicator -ne '=>' } | Select-Object -ExpandProperty InputObject | Sort-Object -Unique | ForEach-Object { $propertyList.Add($_) }
                    $sortedData = $dataFromImport | Select-Object -Property $propertyList | Sort-Object -Property Hostname
                }
                else {
                    $sortedData = $dataFromImport | Sort-Object -Property Hostname
                    $propertyList = [System.Collections.Generic.List[System.String]]::New()
                    $propertyList.Add('Hostname')
                    $sortedData | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -ne 'Hostname' } | Select-Object -ExpandProperty Name | ForEach-Object { $propertyList.Add($_) }
                }
            }

            $xl = New-Object -ComObject Excel.Application
            $wb = $xl.Workbooks.Add()
            $ws = $wb.Sheets(1)

            $firstColumnIndex = 'A'
            $firstRowIndex = '1'
            $lastColumnIndex = $numberToAlphabet[$($sortedData | Get-Member -MemberType NoteProperty).Count]
            $lastRowIndex = $sortedData.Count + 1

            $range = $ws.Range("${firstColumnIndex}${firstRowIndex}:${lastColumnIndex}${lastRowIndex}")
            $rowCount = 1
            $headerRow = $range.Rows.Item($rowCount)
            $headerRow.HorizontalAlignment = 3
            $cellCount = 1

            ForEach ($property in $propertyList) {
                $headerRow.Cells.Item($cellCount).Value = $property
                $cellCount++
            }
            $rowCount++

            ForEach ($system in $sortedData) {
                $currentRow = $range.Rows.Item($rowCount)
                $cellCount = 1
                ForEach ($property in $propertyList) {
                    $currentRow.Cells.Item($cellCount).Value = $system.$property
                    $cellCount++
                }
                $rowCount++
            }

            $sourceType = [Microsoft.Office.Interop.Excel.xlListObjectSourceType]::xlSrcRange
            $xlHeaders = [Microsoft.Office.Interop.Excel.xlYesNoGuess]::xlYes
            $table = $ws.ListObjects.Add($sourceType, $range, $xlHeaders)
            $range.Columns.AutoFit() | Out-Null
            $sortColumn = $ws.Range("${firstColumnIndex}${firstRowIndex}:${firstColumnIndex}${lastRowIndex}")
            $ws.UsedRange.Sort($sortColumn, 1) | Out-Null

            $xlFormat = [Microsoft.Office.Interop.Excel.xlFileFormat]::xlWorkbookDefault

            if (!($ExcelPath)) { $ExcelPath = 'Sorted_Data' }

            $ws.SaveAs($ExcelPath, $xlFormat)
            $wb.Close()
            $xl.Quit()
        }
        'NoImport' {
            if (!(Test-Path -Path $ExcelPath)) {
                Write-Error "File not found at $ExcelPath"
                return
            }

            $xl = New-Object -ComObject Excel.Application
            $wb = $xl.Workbooks.Open($ExcelPath)
            $ws = $wb.Sheets(1)

            $propertyList = [System.Collections.Generic.List[System.String]]::New()
            $headerRow = $ws.UsedRange.Rows(1)
            $cellCount = 1
            do {
                $propertyList.Add($($headerRow.Cells($cellCount).Value2))
                $cellCount++
            } until ($cellCount -gt $headerRow.Cells.Count)

            $xlData = [System.Collections.Generic.List[System.Object]]::New()
            $ws.UsedRange.Rows | Select-Object -Skip 1 | ForEach-Object {
                $currentRow = $_
                $hash = @{}
                $cellCount = 1
                ForEach ($property in $propertyList) {
                    $hash.Add($property, $currentRow.Cells($cellCount).Value2.ToString())
                    $cellCount++
                }
                $xlData.Add($(New-Object -TypeName PSObject -Property $hash))
            }

            $sortedData = $xlData | Select-Object -Property $propertyList | Sort-Object -Property Hostname
            $ws = $wb.Sheets.Add()

            $firstColumnIndex = 'A'
            $firstRowIndex = '1'
            $lastColumnIndex = $numberToAlphabet[$($sortedData | Get-Member -MemberType NoteProperty).Count]
            $lastRowIndex = $sortedData.Count + 1

            $range = $ws.Range("${firstColumnIndex}${firstRowIndex}:${lastColumnIndex}${lastRowIndex}")
            $rowCount = 1
            $headerRow = $range.Rows($rowCount)
            $headerRow.HorizontalAlignment = 3
            $cellCount = 1

            ForEach ($property in $propertyList) {
                $headerRow.Cells($cellCount).Value = $property
                $cellCount++
            }
            $rowCount++

            ForEach ($system in $sortedData) {
                $currentRow = $range.Rows($rowCount)
                $cellCount = 1
                ForEach ($property in $propertyList) {
                    $currentRow.Cells($cellCount).Value = $system.$property
                    $cellCount++
                }
                $rowCount++
            }

            $sourceType = [Microsoft.Office.Interop.Excel.xlListObjectSourceType]::xlSrcRange
            $xlHeaders = [Microsoft.Office.Interop.Excel.xlYesNoGuess]::xlYes
            $table = $ws.ListObjects.Add($sourceType, $range, $xlHeaders)
            $range.Columns.AutoFit() | Out-Null
            $sortColumn = $ws.Range("${firstColumnIndex}${firstRowIndex}:${firstColumnIndex}${lastRowIndex}")
            $ws.UsedRange.Sort($sortColumn, 1) | Out-Null

            $wb.Save()
            $wb.Close()
            $xl.Quit()
        }
    }
}