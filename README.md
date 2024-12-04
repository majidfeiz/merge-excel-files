
# Excel File Merger

A simple Python script to merge multiple Excel files from a specified folder into a single Excel file. This script uses `pandas` and `openpyxl` libraries to handle and process Excel files.

## Features

- Merges all `.xlsx` files in a specified directory.
- Combines data from multiple Excel files into a single output file.
- Handles errors gracefully if a file cannot be read.

## Requirements

- Python 3.7+
- Required Python libraries:
  - `pandas`
  - `openpyxl`

## Installation

1. Clone this repository or download the script.

   ```bash
   git clone https://github.com/yourusername/merge-excel-files.git
   cd merge-excel-files
   ```

2. Install the required Python libraries.

   ```bash
   pip install pandas openpyxl
   ```

## Usage

1. Set the `path` variable in the script to the directory containing the `.xlsx` files you want to merge. Example:

   ```python
   path = "/path/to/your/excel/files/"
   ```

2. Run the script.

   ```bash
   python merge_excel_files.py
   ```

3. The merged Excel file will be saved as `result2.xlsx` in the current working directory.

## Error Handling

- If a file cannot be read, the script will log an error message with the file name and continue processing other files.
- If no valid Excel files are found in the specified directory, the script will print a message and exit.

## Example Directory Structure

```
/path/to/your/excel/files/
    ├── file1.xlsx
    ├── file2.xlsx
    ├── file3.xlsx
```

## Output

After running the script, the merged file `result2.xlsx` will be generated in the script's working directory, containing all the data from the `.xlsx` files in the specified directory.

## Contributing

Feel free to fork this repository and submit pull requests for improvements or additional features.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
