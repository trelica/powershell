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

function AddPropertyIfNotNullOrEmpty($hash, $key, $value) {
    if ($null -ne $value -and $value -ne '') {
        $hash[$key] = $value
    }
}
function Convert-ToCustomFieldRef {
    param (
        [string]$inputString
    )
    
    $lowercaseString = $inputString.ToLower()
    $underscoredString = $lowercaseString -replace ' ', '_'
    # Remove any non-alphanumeric and non-underscore characters
    $cleanString = $underscoredString -replace '[^a-z0-9_]', ''
    $resultString = "user_$cleanString"    
    return $resultString
}

function Convert-AssetStatus {
    param (
        [string]$fromStatus
    )
    
    # Mapping list
    $mapping = @{
        'Deployed'      = 'Deployed'
        'Retired'       = 'Retired'
        'OutForRepair'  = 'OutForRepair'
        'In repair'     = 'OutForRepair'
        'ReadyToDeploy' = 'ReadyToDeploy'
        'Ready'         = 'ReadyToDeploy'
        'New'           = 'New'
    }
    
    # Check if the input status is in the mapping list
    if ($mapping.ContainsKey($fromStatus)) {
        return $mapping[$fromStatus]
    }
    # Call Convert-ToCustomFieldRef for unmapped values
    return Convert-ToCustomFieldRef -inputString $fromStatus
}

function MapFieldToOptionId {
    param (
        [array]$customFields,
        [string]$lookupKey,
        [string]$fieldValue
    )

    $field = $customFields | Where-Object { $_.lookupKey -eq $lookupKey }
    if ($null -eq $field) {
        # just means it's not a lookup field so pass the value through
        return $fieldValue
    }

    $option = $field.options | Where-Object { $_.label -eq $fieldValue }
    if ($null -eq $option) {
        # check if they passed the raw value
        $option = $field.options | Where-Object { $_.id -eq $fieldValue }
        if ($null -eq $option) {
            Write-Error "Error: Option with label '$fieldValue' not found in custom field '$lookupKey'."
            return $null
        }
    }

    return $option.id
}

# stop on first error
$ErrorActionPreference = "Stop"

# Load the CSV file
$data = Import-Csv -Path $CsvFilePath 

# Define the ID property name
$idPropertyName = 'trelicaId'
$assetTagPropertyName = 'assetTag'

# Identify duplicates and extract their IDs
$duplicateIdValues = $data | Where-Object { $_.$idPropertyName -ne $null -and $_.$idPropertyName -ne '' } |
Group-Object -Property $idPropertyName | Where-Object { $_.Count -gt 1 } |
ForEach-Object { $_.Group.$idPropertyName } | Select-Object -Unique

# Identify duplicates based on assetTag and extract their tags
$duplicateAssetTagValues = $data | Where-Object { $_.$assetTagPropertyName -ne $null -and $_.$assetTagPropertyName -ne '' } |
Group-Object -Property $assetTagPropertyName | Where-Object { $_.Count -gt 1 } |
ForEach-Object { $_.Group.$assetTagPropertyName } | Select-Object -Unique

# Filter out duplicates from the original data
$uniqueEntries = $data | Where-Object {
    $duplicateIdValues -notcontains $_.$idPropertyName -and
    $duplicateAssetTagValues -notcontains $_.$assetTagPropertyName
}

# Output the duplicate Trelica IDs and assetTags
Write-Output "Skipping duplicate Trelica IDs: $($duplicateIdValues -join ', ')"
Write-Output "Skipping duplicate Asset Tags: $($duplicateAssetTagValues -join ', ')"

# Read in the custom fields schema so we can map dropdown option labels to ids
$customfieldsSchema = (Invoke-TrelicaRequest -Context $Context -Path "/api/assets/v1/customfields").results
# $user_asset_stored_map = GetLabelToIdMap $customfields.results "user_asset_stored"
# $user_status_map = GetLabelToIdMap $customfields.results "user_status"

# Now, $uniqueEntries contains everything except the duplicates, including entries where the property is null
# Process the unique entries
for ($index = 0; $index -lt $uniqueEntries.Count; $index++) {
    $row = $uniqueEntries[$index]

    $id = $row."$idPropertyName"

    $data = @{
        customFields = @{}
    }

    AddPropertyIfNotNullOrEmpty $data "status" (Convert-AssetStatus -fromStatus $row."status")
    AddPropertyIfNotNullOrEmpty $data "serialNumber"$row."serialNumber"
    AddPropertyIfNotNullOrEmpty $data "assetTag" $row."assetTag"
    AddPropertyIfNotNullOrEmpty $data "purchasedDate" $row."purchasedDate"
    # AddPropertyIfNotNullOrEmpty $data "location" $row."location"
    AddPropertyIfNotNullOrEmpty $data "assignedToEmail" $row."assignedToEmail"
    AddPropertyIfNotNullOrEmpty $data "warrantyExpirationDate" $row."warrantyExpirationDate"
    AddPropertyIfNotNullOrEmpty $data "notes" $row."notes"

    # Add custom fields conditionally, doing a lookup on any dropdowns
    $fieldsWithoutProviderId = $customFieldsSchema | Where-Object { -not $_.PSObject.Properties['providerId'] }
    foreach ($field in $fieldsWithoutProviderId) {
        $lookupKey = $field.lookupKey
        $value = $row.$lookupKey
        if ($null -eq $value -or $value -eq '' -or $lookupKey -eq 'user_model') {
            continue;
        }
        if ("Option" -eq $field.type) {
            $value = MapFieldToOptionId -customFields $customFieldsSchema -lookupKey $lookupKey -fieldValue $value
        }
        if ("Date" -eq $field.type) {
            $date = [datetime]::ParseExact($value, "yyyy-MM-dd", $null)
            $value = $date.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        $data.customFields[$lookupKey] = $value;
    }

    if ($null -eq $id -or $id.Trim() -eq '') {
        # These are read-only
        AddPropertyIfNotNullOrEmpty $data "modelName" $row."modelName"
        AddPropertyIfNotNullOrEmpty $data "assetType" $row."assetType"
        AddPropertyIfNotNullOrEmpty $data "hardwareVendor" $row."hardwareVendor"
        AddPropertyIfNotNullOrEmpty $data "deviceName" $row."deviceName"

        $header = "[$($index+2)]: Inserting Asset serial number $($row.serialNumber)"
        Write-Output $header
        Write-Output ("=" * ($header.Length))
        $json = ($data | ConvertTo-Json -Depth 6);
        Write-Output "* PUT: $json"
        Invoke-TrelicaRequest -Context $Context -Path "/api/assets/v1" -Method "PUT" -PostData $json
        # Write-Output "* Result: $($asset | ConvertTo-Json -Depth 6)"
    }
    else {
        $header = "[$($index+2)]: Updating Asset ID $id"
        Write-Output $header
        Write-Output ("=" * ($header.Length))
        $json = ($data | ConvertTo-Json -Depth 6);
        Write-Output "* PATCH: $json"
        Invoke-TrelicaRequest -Context $Context -Path "/api/assets/v1/$($id)" -Method "PATCH" -PostData $json
        # Write-Output "* Result: $($asset | ConvertTo-Json -Depth 6)"
    }
}