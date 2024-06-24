# Ensure the ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Import the module
Import-Module ImportExcel

# Define the file path
$filePath = "C:\your-file-path.csv"
$outputDir = "C:\your-output-directory"

# Import the CSV file
$data = Import-Csv -Path $filePath

# Get unique safe owners
$safeOwners = $data | Select-Object -ExpandProperty SafeOwnerName | Sort-Object -Unique

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
        Set-CellStyle -WorkSheet $worksheet -Row $row -LastColumn 'N' -Pattern 'Solid' -Color ([System.Drawing.Color]::FromArgb(255, 255, 199, 206))  # LightSalmon color
    }
    Close-ExcelPackage $excelPackage
}

# Loop through each safe owner and create or update Excel files
foreach ($owner in $safeOwners) {
    # Filter the data for the current safe owner
    $ownerData = $data | Where-Object { $_.SafeOwnerName -eq $owner }

    # Create the new file path
    $newFilePath = Join-Path -Path $outputDir -ChildPath "$($owner)_RemediationList.xlsx"

    # Check if the file already exists
    if (Test-Path $newFilePath) {
        # Import existing data
        $existingData = Import-Excel -Path $newFilePath

        # Find new rows to add based on unique identifiers
        $existingIds = $existingData | Select-Object -ExpandProperty CisarID
        $newRows = $ownerData | Where-Object {
            ($_."CisarID" -ne "" -and $existingIds -notcontains $_."CisarID") -or
            ($_."CisarID" -eq "" -and $existingData | Where-Object { $_."SystemAddress" -eq $_."SystemAddress" } -eq $null)
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

            # Notify that the safe owner's file has been updated
            Write-Output "Updated file for Safe Owner: $owner"
        }
    } else {
        # Insert the "ServiceNow INC reference" column at the beginning
        $ownerData = $ownerData | ForEach-Object {
            $newObject = [PSCustomObject]@{"ServiceNow INC reference" = ""}
            foreach ($property in $_.PSObject.Properties) {
                $newObject | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
            }
            $newObject
        }

        # Export the filtered data to a new Excel file with autofit and table formatting
        $ownerData | Export-Excel -Path $newFilePath -AutoSize -TableName "RemediationList" -TableStyle Light15

        # Apply "bad" cell style to new rows
        $rowCount = ($ownerData | Measure-Object).Count
        $rowStart = 2
        $rowEnd = $rowStart + $rowCount - 1
        Apply-BadStyle -filePath $newFilePath -rowStart $rowStart -rowEnd $rowEnd

        # Notify that the safe owner's file has been created
        Write-Output "Created file for Safe Owner: $owner"
    }
}

# Output completion message
Write-Output "Writing completed"
