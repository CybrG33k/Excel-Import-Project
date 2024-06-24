# Remediation List Automation

This PowerShell script automates the creation and updating of Excel files based on a remediation list CSV file. The script filters the data by unique safe owners, creates new Excel files for each owner, updates existing files, and applies a specific cell style to new rows. It also logs the updates and creations to a text file.

## Features

- Filters data by unique safe owners.
- Creates new Excel files for each safe owner.
- Updates existing Excel files with new rows if needed.
- Applies a "bad" cell style with a lighter red color to new rows.
- Logs updates and creations to a text file.

## Requirements

- PowerShell
- ImportExcel module

## Installation

1. Ensure the `ImportExcel` module is installed:
    ```powershell
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
    ```

2. Download or clone this repository to your local machine.

## Usage

1. Modify the script paths in `process_csv.ps1` to match your file locations:
    - Update the `filePath` variable with the path to your CSV file.
    - Update the `outputDir` variable with the path where you want to save the new or updated Excel files.
    - Modify the `logFilePath` variable to point to where you want to save the log file.

2. Run the script from PowerShell ISE or directly from PowerShell by navigating to the script's directory and executing:
    ```powershell
    .\process_csv.ps1
    ```

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please fork this repository and submit pull requests.

## Contact

If you have any questions, feel free to open an issue or contact the project maintainers.
