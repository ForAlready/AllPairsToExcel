# AllPairsToExcel
使用Allpairs工具对Excel文件中的数据进行组合测试用例生成，并将结果保存到新的工作表中。
# Allpairs-Excel 测试用例生成器

## 项目描述

该项目提供了一个 PowerShell 脚本，使用 Allpairs 工具从 Excel 数据中生成组合测试用例。它有助于自动化测试用例生成过程，提高需要组合测试场景的测试效率。

## 功能

- **自动化测试用例生成**：利用 Allpairs 工具从 Excel 数据中创建高效的测试用例组合。
- **PowerShell 集成**：提供与 PowerShell 5 和 PowerShell 7 兼容的脚本，以增加灵活性。
- **易于使用**：简单的命令行操作，具有清晰的参数要求。

## 入门指南

### 先决条件

- **Allpairs 工具**：确保 `allpairs.exe` 可用在脚本目录或系统路径中。
- **PowerShell**：与 PowerShell 5 和 PowerShell 7 兼容。

### 安装

1. 将此仓库克隆到本地计算机：
   ```bash
   git clone https://github.com/ForAlready/AllPairsToExcel.git
   ```
2. 进入项目目录：
   ```bash
   cd Allpairs-Excel-TestCase-Generator
   ```

## 使用方法

### 对于 PowerShell 5

```powershell
.\ExcelAllpairsPS5.ps1 -ExcelFile "path_to_your_excel_file.xlsx"
```

### 对于 PowerShell 7

```powershell
pwsh ./AllpairsExcelPS7.ps1 -ExcelFile "path_to_your_excel_file.xlsx"
```

## 示例

假设当前目录下有一个名为 `testdata.xlsx` 的 Excel 文件。运行以下命令：

```powershell
.\ExcelAllpairsPS5.ps1 -ExcelFile ".\testdata.xlsx"
```

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

# Allpairs-Excel-TestCase-Generator

## Project Description

This project provides a PowerShell script to generate combination test cases from Excel data using the Allpairs tool. It helps automate the test case generation process, improving testing efficiency for scenarios requiring combinatorial testing.

## Features

- **Automated Test Case Generation**: Utilizes the Allpairs tool to create efficient test case combinations from Excel data.
- **PowerShell Integration**: Offers scripts compatible with both PowerShell 5 and PowerShell 7 for flexibility.
- **Easy to Use**: Simple command-line operation with clear parameter requirements.

## Getting Started

### Prerequisites

- **Allpairs Tool**: Ensure `allpairs.exe` is available in the script directory or system path.
- **PowerShell**: Compatible with PowerShell 5 and PowerShell 7.

### Installation

1. Clone this repository to your local machine:
   ```bash
   git clone https://github.com/ForAlready/AllPairsToExcel.git
   ```
2. Navigate to the project directory:
   ```bash
   cd Allpairs-Excel-TestCase-Generator
   ```

## Usage

### For PowerShell 5

```powershell
.\ExcelAllpairsPS5.ps1 -ExcelFile "path_to_your_excel_file.xlsx"
```

### For PowerShell 7

```powershell
pwsh ./AllpairsExcelPS7.ps1 -ExcelFile "path_to_your_excel_file.xlsx"
```

## Example

Suppose you have an Excel file named `testdata.xlsx` in the current directory. Run the following command:

```powershell
.\ExcelAllpairsPS5.ps1 -ExcelFile ".\testdata.xlsx"
```

This will generate test case combinations and save the results to a new worksheet named "AllpairsResults" in the `testdata.xlsx` file.

## Notes

- The script processes the first worksheet of the Excel file.
- A temporary CSV file is created during processing and automatically deleted afterward.
- If the "AllpairsResults" worksheet already exists, it may be overwritten.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Special thanks to the creators of the Allpairs tool for providing the foundation for this test case generation solution.



