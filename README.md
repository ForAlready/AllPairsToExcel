# AllPairsToExcel

## 项目描述

AllPairsToExcel 是一个提供 PowerShell 脚本的项目，使用 Allpairs 工具从 Excel 数据中生成组合测试用例。此工具有助于自动化测试用例生成过程，提高需要组合测试场景的测试效率。

## 功能

- **自动化测试用例生成**：利用 Allpairs 工具从 Excel 数据中创建高效的测试用例组合。
- **PowerShell 集成**：提供与 PowerShell 5 （windows系统自带）和 PowerShell 7 兼容的脚本，以增加灵活性。
- **易于使用**：简单的命令行操作，具有清晰的参数要求。
- **拖放支持**：允许将 Excel 文件拖放到批处理文件上，轻松执行脚本。

## 入门指南

### 先决条件

- **Allpairs 工具**：确保 `allpairs.exe`和`allpairs.pl` 与AllpairsExcelPS5.ps1/AllpairsExcelPS7.ps1在同一目录或在系统环境变量（Path）中。
- **PowerShell**：与 PowerShell 5 和 PowerShell 7 兼容。

### 安装

1. 将此仓库克隆到本地计算机：
   ```bash
   git clone https://github.com/ForAlready/AllPairsToExcel.git
   ```
2. 进入项目目录：
   ```bash
   cd AllPairsToExcel
   ```
3. 解压allpairs.zip将压缩包中的 `allpairs.exe`和`allpairs.pl`移动到与AllpairsExcelPS5.ps1/AllpairsExcelPS7.ps1在同一目录中

## 使用方法

### 对于 PowerShell 5

#### 使用命令行
```powershell
.\ExcelAllpairsPS5.ps1 -ExcelFile "path_to_your_excel_file.xlsx"
```

#### 使用拖放
1. 找到项目目录中的 `powershell5.bat` 文件。
2. 将您的 Excel 文件拖放到 `powershell5.bat` 文件上。
3. 这将自动以您的 Excel 文件作为输入执行 PowerShell 5 脚本。

### 对于 PowerShell 7

#### 使用命令行
```powershell
pwsh ./AllpairsExcelPS7.ps1 -ExcelFile "path_to_your_excel_file.xlsx"
```

#### 使用拖放
1. 找到项目目录中的 `pwsh.bat` 文件。
2. 将您的 Excel 文件拖放到 `pwsh.bat` 文件上。
3. 这将自动以您的 Excel 文件作为输入执行 PowerShell 7 脚本。

## 示例

假设当前目录下有一个名为 `testdata.xlsx` 的 Excel 文件。运行以下命令：

```powershell
.\ExcelAllpairsPS5.ps1 -ExcelFile ".\testdata.xlsx"
```

或者将 `testdata.xlsx` 拖放到 `powershell5.bat` 上。

这将生成测试用例组合，并将结果保存到 `testdata.xlsx` 文件的新工作表 "AllpairsResults" 中。

## 注意事项

- 该脚本处理 Excel 文件的第一个工作表。
- 在处理过程中会创建一个临时的 CSV 文件，处理完成后会自动删除。
- 如果 "AllpairsResults" 工作表已存在，可能会被覆盖。

## 许可证

此项目根据 MIT 许可证授权 - 详情请参见 LICENSE 文件。

## 致谢

- 感谢 Allpairs 工具的创建者为这个测试用例生成解决方案提供了基础。

<hr>

# AllPairsToExcel

## Project Description

AllPairsToExcel is a project that provides a PowerShell script to generate combination test cases from Excel data using the Allpairs tool. This tool helps automate the test case generation process, improving testing efficiency for scenarios requiring combinatorial testing.

## Features

- **Automated Test Case Generation**: Utilizes the Allpairs tool to create efficient test case combinations from Excel data.
- **PowerShell Integration**: Offers scripts compatible with both PowerShell 5（It comes pre-installed with the Windows operating system. ） and PowerShell 7 for flexibility.
- **Easy to Use**: Simple command-line operation with clear parameter requirements.
- **Drag-and-Drop Support**: Allows you to drag and drop Excel files onto batch files to easily execute the script.

## Getting Started

### Prerequisites

- **Allpairs Tool**: Ensure `allpairs.exe` and `allpairs.pl` are available in the script directory or system path.
- **PowerShell**: Compatible with PowerShell 5 and PowerShell 7.

### Installation

1. Clone this repository to your local machine:
   ```bash
   git clone https://github.com/ForAlready/AllPairsToExcel.git
   ```
2. Navigate to the project directory:
   ```bash
   cd AllPairsToExcel
   ```
3.Unzip the allpairs.zip file, and move the `allpairs.exe` and `allpairs.pl` in the compressed package to the same directory as AllpairsExcelPS5.ps1/AllpairsExcelPS7.ps1. 

## Usage

### For PowerShell 5

#### Using Command Line
```powershell
.\ExcelAllpairsPS5.ps1 -ExcelFile "path_to_your_excel_file.xlsx"
```

#### Using Drag-and-Drop
1. Locate the `powershell5.bat` file in the project directory.
2. Drag your Excel file and drop it onto the `powershell5.bat` file.
3. This will automatically execute the PowerShell 5 script with your Excel file as the input.

### For PowerShell 7

#### Using Command Line
```powershell
pwsh ./AllpairsExcelPS7.ps1 -ExcelFile "path_to_your_excel_file.xlsx"
```

#### Using Drag-and-Drop
1. Locate the `pwsh.bat` file in the project directory.
2. Drag your Excel file and drop it onto the `pwsh.bat` file.
3. This will automatically execute the PowerShell 7 script with your Excel file as the input.

## Example

Suppose you have an Excel file named `testdata.xlsx` in the current directory. Run the following command:

```powershell
.\ExcelAllpairsPS5.ps1 -ExcelFile ".\testdata.xlsx"
```

Or drag and drop `testdata.xlsx` onto `powershell5.bat`.

This will generate test case combinations and save the results to a new worksheet named "AllpairsResults" in the `testdata.xlsx` file.

## Notes

- The script processes the first worksheet of the Excel file.
- A temporary CSV file is created during processing and automatically deleted afterward.
- If the "AllpairsResults" worksheet already exists, it may be overwritten.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Special thanks to the creators of the Allpairs tool for providing the foundation for this test case generation solution.
