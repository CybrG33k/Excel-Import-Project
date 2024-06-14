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

# Loop through each safe owner and create a new Excel file
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

    # Export the filtered data to a new Excel file with autofit and table formatting
    $ownerDataArray | Export-Excel -Path $newFilePath -AutoSize -TableName "RemediationList" -TableStyle Light15
}

# Output completion message
Write-Output "Writing completed"
