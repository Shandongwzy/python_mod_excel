# Excel Processor

`excel_processor.exe` is a standalone tool for automating modifications to Excel files (.xls or .xlsx) based on user-defined rules. It reads a specified column in an input Excel file, applies regular expression (regex) matching rules defined in `rules.xls`, and updates designated columns in an output Excel file. The tool preserves formatting in `.xls` files and supports Unicode (e.g., Chinese characters) in file names and cell content. This is ideal for automating repetitive data updates, such as filling specific values based on text patterns.

**Important**: Always back up all Excel files in the working directory before running the program to prevent data loss.

## Features
- **Automated Excel Editing**: Matches text in a specified column using regex and updates other columns with predefined values.
- **Format Preservation**: Retains cell formatting (e.g., fonts, colors) for `.xls` files.
- **Flexible Input/Output**: Supports both `.xls` and `.xlsx` files, with input and output files configurable in `rules.xls`.
- **Unicode Support**: Handles non-ASCII characters (e.g., Chinese) in file names and cell content.
- **Extensible Rules**: Allows multiple column updates per rule, defined in `rules.xls`.

## Prerequisites
- **Windows OS**: The executable is built for Windows.
- **Excel Files**: Input and output files must be in `.xls` or `.xlsx` format.
- **No Software Installation Required**: The executable runs standalone, but `rules.xls` must be configured correctly.

## Setup
1. **Download the Executable**: Obtain `excel_processor.exe` from the [Releases](https://github.com/your-repo/excel_processor/releases) page.
2. **Prepare Files**: Place the following in the same directory as `excel_processor.exe`:
   - `rules.xls`: Configuration file defining matching and update rules (must be named exactly `rules.xls`).
   - Input Excel file (e.g., `input.xls` or `input.xlsx`).
   - Output Excel file (can be the same as the input file or different).
3. **Backup Files**: Save copies of all Excel files to prevent accidental data loss.

## Configuring `rules.xls`
The `rules.xls` file defines how the program processes Excel files. Create or edit it using Microsoft Excel, ensuring the first row contains specific column headers (case-sensitive) and subsequent rows define rules.

### Required Columns
| Column Name      | Description                                                                 |
|------------------|-----------------------------------------------------------------------------|
| Input_File       | Name of the input Excel file (e.g., `input.xls`).                           |
| Input_Sheet      | Sheet name (e.g., `Sheet1`) or index (e.g., `0` for the first sheet).       |
| Regex            | Regular expression to match text in the specified column (e.g., `.*terminal.*`). |
| Regex_Column     | Column to check for regex matches (e.g., `H`, `7`, or header name like `Description`). |
| Output_File      | Name of the output Excel file (e.g., `output.xls`; can match Input_File).   |
| Output_Sheet     | Sheet name or index for output (e.g., `Sheet1` or `0`).                     |
| Change1_Column   | Column to update if regex matches (e.g., `AX` or header name).              |
| Change1_Value    | Value to write to Change1_Column (e.g., `27`).                              |
| Change2_Column   | Optional second column to update (e.g., `K`).                               |
| Change2_Value    | Optional value for Change2_Column (e.g., `123`).                            |

### Optional Columns
- Add more pairs like `Change3_Column`/`Change3_Value`, `Change4_Column`/`Change4_Value`, etc., for additional updates per rule.

### Example `rules.xls`
| Input_File   | Input_Sheet | Regex           | Regex_Column | Output_File  | Output_Sheet | Change1_Column | Change1_Value | Change2_Column | Change2_Value |
|--------------|-------------|-----------------|--------------|--------------|--------------|----------------|---------------|----------------|---------------|
| input.xls    | Sheet1      | .*terminal.*    | H            | output.xls   | Sheet1       | AX             | 27            |                |               |
| input.xls    | 0           | .*breaker.*     | H            | input.xls    | 0            | AX             | 27            | K              | 123           |

### Notes
- **File Names**: Ensure `Input_File` and `Output_File` match the actual file names in the directory, including extensions.
- **Sheet Names**: Sheet names are case-sensitive. Use `0` for the first sheet if unsure.
- **Regex**: Use valid regex patterns (see below). Simple patterns like `*terminal*` are not supported; use `.*terminal.*` instead.
- **Column Identifiers**: Specify columns by letter (e.g., `H`), index (e.g., `7`), or header name (e.g., `Description`).

## Running the Program
1. **Open Command Prompt**:
   - Navigate to the directory containing `excel_processor.exe` (e.g., right-click the folder, select "Open in Terminal" or "Command Prompt").
2. **Execute the Program**:
   - excel_processor.exe
3. **Check Output**:
- The output file will be updated based on `rules.xls`.
- A `log.txt` file is generated in the same directory, logging success (e.g., "Saved .xls to ...") or errors (e.g., invalid regex or missing files).
4. **Troubleshooting**:
- If no changes are applied, verify `rules.xls` (correct regex, column names, file paths).
- Check `log.txt` for error details.
- Ensure input file has data in the specified `Regex_Column`.

## Regular Expression Guide
Regular expressions (regex) define text patterns to match in the `Regex_Column`. The program requires valid Python regex syntax.

### Simplified Usage
- **Basic Matching**: Use literal text (e.g., `terminal` matches "terminal" exactly).
- **Wildcard Equivalent**: Use `.*text.*` to match any text containing "text" (e.g., `.*terminal.*` matches "abc terminal def").
- **Test Patterns**: Use tools like [regex101.com](https://regex101.com/) (select Python flavor) to validate patterns.

### Common Patterns
| Pattern         | Description                                    |
|-----------------|------------------------------------------------|
| `text`          | Matches exact text (e.g., `terminal`).         |
| `.*text.*`      | Matches any string containing "text".          |
| `^text`         | Matches strings starting with "text".          |
| `text$`         | Matches strings ending with "text".           |
| `text|other`    | Matches "text" or "other".                    |
| `[0-9]`         | Matches any single digit.                     |
| `[a-zA-Z]`      | Matches any single letter.                    |

### Tips
- **Avoid Invalid Patterns**: Do not use standalone `*` (e.g., `*terminal*` is invalid; use `.*terminal.*`).
- **Escape Special Characters**: To match `.`, `*`, `?`, etc., prefix with `\` (e.g., `\.` matches a literal dot).
- **Unicode Support**: Chinese characters (e.g., `端子`) are treated as literal characters and work directly.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
