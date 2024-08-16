[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 1)][String]$CsvFilePath,
    [Parameter(Mandatory = $true, Position = 2)][String]$Secret,
    [Parameter(Mandatory = $true, Position = 3)][String]$Url
)
<#
    .SYNOPSIS
    Runs a workflow for all elements in a CSV file  
    .EXAMPLE
    /Run-Workflow.ps1 -CsvFilePath '~/tmp/assets.csv' -Secret 'XXXXXX' -Url 'https://dev.trelica.com/FlowApi/webhooks/66befd15027db4366b802065/w6o4ffUyxXBjKLYZa0cSLy01Esm'
#>

# stop on first error
$ErrorActionPreference = "Stop"

# Load the CSV file
$csvData = Import-Csv -Path $CsvFilePath 

# Convert the CSV data to the desired JSON structure
$jsonObject = @{
    items = $csvData
}

# Convert the object to JSON format
$json = $jsonObject | ConvertTo-Json -Depth 3

Write-Output "POST: $json"
# Set the headers including the x-secret header
$headers = @{
    "x-secret"     = $Secret
    "Content-Type" = "application/json"
}

# POST the JSON data to the specified URL
$response = Invoke-RestMethod -Uri $Url -Method POST -Body $json -Headers $headers
Write-Output $response
