<#
.SYNOPSIS
使用Allpairs工具对Excel文件中的数据进行组合测试用例生成，并将结果保存到新的工作表中。

.DESCRIPTION
该脚本需使用powershell 7 或更高的版本,需使用接受一个Excel文件作为输入，读取其中的数据，使用Allpairs工具进行组合测试用例生成，然后将生成的测试用例结果写入到Excel文件的新工作表中。该脚本适用于需要进行组合测试的场景，能够自动生成测试用例组合，提高测试效率。

.PARAMETER ExcelFile
指定要处理的Excel文件的路径。该参数是必须的，且需要位于参数列表的第一个位置。脚本会检查该文件是否存在，如果不存在将报错并退出。

.EXAMPLE
PS> .\AllpairsToExcel.ps1 .\testdata.xlsx

说明：将当前目录下的testdata.xlsx文件作为输入，使用Allpairs工具生成测试用例组合，并将结果保存到testdata.xlsx文件的新工作表"AllpairsResults"中。

.NOTES
- 该脚本需使用powershell 7 或更高的版本（pwsh）。
- 脚本需要在安装了Allpairs工具的环境中运行，并且allpairs.exe需要位于脚本所在目录或系统路径中。
- 脚本处理的是Excel文件的第一个工作表。
- 生成的测试用例结果会保存到新的工作表"AllpairsResults"中，如果该工作表已存在，可能会被覆盖或导致错误。
- 脚本在处理过程中会创建一个临时的CSV文件，处理完成后会自动删除该临时文件。
- 感谢作者“小可乐”为本脚本编写了该脚本和详细的说明文档。

.INPUTS
System.String. 即Excel文件的路径(支持相对路径)。

.OUTPUTS
无直接输出，结果会保存到指定的Excel文件的新工作表中。

.LINK
Allpairs工具官方网站或相关文档
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
$data | Export-Csv -Path $tempCSVFile -Delimiter "`t" -NoTypeInformation -Encoding ANSI

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