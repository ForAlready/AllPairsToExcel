<#
.SYNOPSIS
Generates combination test cases from Excel data using Allpairs tool and saves results to a new worksheet.

.DESCRIPTION
This script takes an Excel file as input, reads its data, generates combination test cases using the Allpairs tool, and writes the results to a new worksheet in the Excel file. It's designed for combination testing scenarios, automating test case generation to improve testing efficiency.

.PARAMETER ExcelFile
Specifies the path of the Excel file to process. This mandatory parameter must be the first in the list. The script checks if the file exists, throwing an error and exiting if it doesn't.

.EXAMPLE
PS> .\AllpairsToExcel.ps1 .\testdata.xlsx

This example uses the testdata.xlsx file from the current directory as input, generates test case combinations with Allpairs, and saves the results to the new "AllpairsResults" worksheet in the testdata.xlsx file.

.NOTES
- This script requires PowerShell 5.
- The script requires an environment with Allpairs installed, and allpairs.exe must be in the script directory or system path.
- It processes the first worksheet of the Excel file.
- Test case results are saved to the new "AllpairsResults" worksheet, which may be overwritten or cause errors if it already exists.
- A temporary CSV file is created during processing and automatically deleted afterward.
- Thank the author "xkl" for writing this script and the detailed instruction document for it. 

.INPUTS
System.String. The path to the Excel file(relative paths are supported).

.OUTPUTS
No direct output; results are saved to a new worksheet in the specified Excel file.

.LINK
- https://github.com/ForAlready/AllPairsToExcel
- Allpairs official website or relevant documentation
#>
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$ExcelFile
)

# 将相对路径转换为绝对路径
$ExcelFile = Resolve-Path $ExcelFile

# 提取输入文件的名称和路径
$ExcelFileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelFile)

# 检查输入文件是否存在
if (-not (Test-Path $ExcelFile)) {
    Write-Error "Excel 文件不存在: $ExcelFile"
    exit 1
}

# 读取 Excel 文件
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($ExcelFile)
$worksheet = $workbook.Sheets.Item(1)  # 假设只处理第一个工作表

# 获取工作表的数据
$rowMax = $worksheet.UsedRange.Rows.Count
$colMax = $worksheet.UsedRange.Columns.Count

# 获取表头并保存列的顺序
$headers = @()
for ($col = 1; $col -le $colMax; $col++) {
    $headers += $worksheet.Cells.Item(1, $col).Text
}

# 创建一个数组来存储数据
$data = @()

# 遍历工作表的数据
for ($row = 2; $row -le $rowMax; $row++) {
    $rowData = @{}
    for ($col = 1; $col -le $colMax; $col++) {
        $value = $worksheet.Cells.Item($row, $col).Text
        if ([string]::IsNullOrEmpty($value)) {
            $value = $null  # 将空值设置为 $null
        }
        $header = $headers[$col - 1]
        $rowData[$header] = $value
    }
    $data += New-Object PSObject -Property $rowData
}

# 将数据导出为临时 CSV 文件，并指定分隔符
$tempCSVFile = "$ExcelFileName_temp.csv"
# $data | Export-Csv -Path $tempCSVFile -Delimiter "`t" -NoTypeInformation -Encoding UTF8
$data | Export-Csv -Path $tempCSVFile -Delimiter "`t" -NoTypeInformation -Encoding Default

# 将 allpairs 的结果编码转换为 ANSI 并写入原本的 CSV 文件
#Get-Content $tempCSVFile | Set-Content -Path $tempCSVFile -Encoding Default·

# 调用 allpairs.exe
$allpairsOutput = & .\allpairs.exe $tempCSVFile


# 创建一个新的工作表
$newWorksheet = $workbook.Sheets.Add()
$newWorksheet.Name = "AllpairsResults"

# 将 allpairs 的输出写入新工作表
$rowIndex = 1
foreach ($line in $allpairsOutput) {
    $colIndex = 1
    $values = $line -split "`t"
    foreach ($value in $values) {
        # 移除所有引号
        $cleanValue = $value -replace '"', ''
        $newWorksheet.Cells.Item($rowIndex, $colIndex) = $cleanValue
        $colIndex++
    }
    $rowIndex++
}

# 调整列宽
$newWorksheet.Columns.AutoFit()

# 保存工作簿
$workbook.Save()

# 关闭 Excel
$excel.Quit()

# 删除临时 CSV 文件
Remove-Item $tempCSVFile

Write-Host "转换完成，新工作表已添加到: $ExcelFile"
