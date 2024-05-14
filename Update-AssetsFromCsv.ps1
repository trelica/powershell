[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)][PSTypeName('Trelica.Context')]$Context,
    [Parameter(Mandatory = $true, Position = 2)][String]$CsvFilePath
)
<#
    .SYNOPSIS
    Applies changes in a CSV file to assets.
    .EXAMPLE
    Run this passing in a Trelica Context:
    Initialize-TrelicaCredentials | Get-TrelicaContext | ./Update-AssetsFromCsv.ps1 -CsvFilePath '~/tmp/assets.csv'
#>

# stop on first error
$ErrorActionPreference = "Stop"

# Load the CSV file
$data = Import-Csv -Path $CsvFilePath 

# Group data by the "Trelica ID" column and filter groups with a count greater than one
$duplicates = $data | Group-Object -Property 'Trelica ID' | Where-Object { $_.Count -gt 1 }

# Output the duplicate Trelica IDs
$duplicates | ForEach-Object {
    Write-Output "Skipping duplicate Trelica ID: $($_.Name)"
}

# Group data by the "Trelica ID" column and filter groups with a count equal to one
$uniqueEntries = $data | Group-Object -Property 'Trelica ID' | Where-Object { $_.Count -eq 1 }

# Read in the custom fields schema so we can map dropdown option labels to ids
$customfields = (Invoke-TrelicaRequest -Context $Context -Path "/api/assets/v1/customfields" | ConvertFrom-Json)
$user_asset_stored_map = GetLabelToIdMap $customfields.results "user_asset_stored"
$user_status_map = GetLabelToIdMap $customfields.results "user_status"

# Process the unique entries
$uniqueEntries | ForEach-Object {
    $row = ($_.Group | Select-Object -First 1)  # Process only the first entry in each group (if multiple)

    $id = $row."Trelica ID"

    $data = [PSCustomObject]@{
        customFields = @{
            user_asset_stored = $user_asset_stored_map[$row."In Office Location"]
            user_status       = $user_status_map[$row."Device Status"]
        }
    }
    $json = ($data | ConvertTo-Json -Depth 6);
    $header = "Updating Asset ID $id"
    Write-Output $header
    Write-Output ("=" * ($header.Length))
    Write-Output "* PATCH: $json"
    $asset = Invoke-TrelicaRequest -Context $Context -Path "/api/assets/v1/$($id)" -Method "PATCH" -PostData $json
    Write-Output "* Result: $($asset | ConvertTo-Json -Depth 6)"
}

function GetLabelToIdMap($results, $fieldName) {
    $dropdownOptions = $results | Where-Object { $_.lookupKey -eq $fieldName }

    $labelToIdMap = @{}
    foreach ($option in $dropdownOptions.options) {
        $labelToIdMap[$option.label] = $option.id
    }

    return $labelToIdMap
}
