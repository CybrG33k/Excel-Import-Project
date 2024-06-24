# Ensure the ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Import the module
Import-Module ImportExcel

# Define the file path
$filePath = "C:\your-file-path.csv"
$outputDir = "C:\your-output-directory"
$logFilePath = "D:\your-update_log-directory.txt"

# Function to apply "Bad" style
function Apply-BadStyle {
    param (
        [string]$filePath,
        [int]$rowStart,
        [int]$rowEnd
    )

    $excelPackage = Open-ExcelPackage -Path $filePath
    $worksheet = $excelPackage.Workbook.Worksheets["Sheet1"]
    for ($row = $rowStart; $row -le $rowEnd; $row++) {
        Set-CellStyle -WorkSheet $worksheet -Row $row -LastColumn 'Z' -Pattern 'Solid' -Color ([System.Drawing.Color]::FromArgb(255, 240, 128, 128))  # LightSalmon color
    }
    Close-ExcelPackage $excelPackage
}

# Import the CSV file
$data = Import-Csv -Path $filePath

# Get unique safe owners
$safeOwners = $data | Select-Object -ExpandProperty SafeOwnerName | Sort-Object -Unique

# Clear the log file if it exists
if (Test-Path $logFilePath) {
    Remove-Item $logFilePath
}

# Loop through each safe owner and create or update Excel files
foreach ($owner in $safeOwners) {
    # Filter the data for the current safe owner
    $ownerData = $data | Where-Object { $_.SafeOwnerName -eq $owner }

    # Create a new array to hold modified objects with the new column
    $ownerDataArray = @()

    foreach ($row in $ownerData) {
        # Create a new custom object with "ServiceNow reference" as the first column
        $newRow = [PSCustomObject]@{
            "ServiceNow reference" = ""
        }
        
        # Add the existing columns from the original row
        $row.PSObject.Properties | ForEach-Object {
            Add-Member -InputObject $newRow -MemberType NoteProperty -Name $_.Name -Value $_.Value
        }

        # Add the modified row to the array
        $ownerDataArray += $newRow
    }

    # Create the new file path
    $newFilePath = Join-Path -Path $outputDir -ChildPath "$($owner)_RemediationList.xlsx"

    # Check if the file already exists
    if (Test-Path $newFilePath) {
        # Import existing data
        $existingData = Import-Excel -Path $newFilePath

        # Find new rows to add based on unique identifiers
        $existingIds = $existingData | Select-Object -ExpandProperty CisarID
        $existingAddresses = $existingData | Select-Object -ExpandProperty SystemAddress
        $newRows = $ownerDataArray | Where-Object {
            ($_."CisarID" -ne "" -and $existingIds -notcontains $_."CisarID") -or
            ($_."CisarID" -eq "" -and $_."SystemAddress" -ne "" -and $existingAddresses -notcontains $_."SystemAddress")
        }

        if ($newRows) {
            # Add new rows to existing data
            $updatedData = @()
            $updatedData += $existingData
            $updatedData += $newRows

            # Export the updated data to the Excel file
            $updatedData | Export-Excel -Path $newFilePath -AutoSize -TableName "RemediationList" -TableStyle Light15

            # Apply "bad" cell style to new rows
            $newRowsCount = ($newRows | Measure-Object).Count
            $rowStart = ($existingData | Measure-Object).Count + 2
            $rowEnd = $rowStart + $newRowsCount - 1
            Apply-BadStyle -filePath $newFilePath -rowStart $rowStart -rowEnd $rowEnd

            # Log the update
            Add-Content -Path $logFilePath -Value "Updated file for Safe Owner: $owner"
        }
    } else {
        # Export the filtered data to a new Excel file with autofit and table formatting
        $ownerDataArray | Export-Excel -Path $newFilePath -AutoSize -TableName "RemediationList" -TableStyle Light15

        # Apply "bad" cell style to new rows
        $rowCount = ($ownerDataArray | Measure-Object).Count
        $rowStart = 2
        $rowEnd = $rowStart + $rowCount - 1
        Apply-BadStyle -filePath $newFilePath -rowStart $rowStart -rowEnd $rowEnd

        # Log the creation
        Add-Content -Path $logFilePath -Value "Created file for Safe Owner: $owner"
    }
}

# Output completion message
Write-Output "Writing completed"
Add-Content -Path $logFilePath -Value "Writing completed"
