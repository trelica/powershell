[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)][PSTypeName('Trelica.Context')]$Context,
    [Parameter(Mandatory = $true, Position = 2)][String]$Filter
)
<#
    .SYNOPSIS
    Exports people and associated app data to CSV, based on a filter.
    .EXAMPLE
    Run this passing in a Trelica Context:
    Initialize-TrelicaCredentials | Get-TrelicaContext | .\Get-AppsForPeople.ps1 -Filter 'email sw "pbs-"' | Export-Csv -Path .\appsForPeople.csv
#>
Function ProcessAppsForPerson() {
    Param(
        [Parameter(Mandatory = $true)][System.Object]$Person,
        [Parameter(Mandatory = $true)][System.Object]$Apps
    )
    If ($Apps.Length -eq 0) {
        # always emit something
        [pscustomobject] [ordered] @{
            'Email' = $person.email; 
            'App Name' = $null;
            'App Status' = $null; 
            'App Id' = $null; 
            'Last Login' = $null;
        }
    } Else {
        Foreach ($app in $Apps) {
            [pscustomobject] [ordered] @{
                'Email' = $person.email; 
                'App Name' = $app.name;
                'App Status' = $app.status; 
                'App Id' = $app.id;
                'Last Login' = $app.appUser.lastLoginDtm
            }
        }
    }
}

Function ProcessPeople() {
    Param(
        [Parameter(Mandatory = $true)][PSTypeName('Trelica.Context')]$Context,
        [Parameter(Mandatory = $true)][System.Object]$People
    )
    Foreach ($person in $People) {
        $filter = 'appUser.status eq "Active"'
        $path = "/api/people/v1/$($person.id)/apps?filter=$([System.Web.HttpUtility]::UrlEncode($filter))"
        Do {
            $apps = Invoke-TrelicaRequest -Context $Context -Path $path
            ProcessAppsForPerson -Person $person -Apps $apps.results
            $path = $apps.next
        } While ($path)
    }
}

$path = "/api/people/v1?filter=$([System.Web.HttpUtility]::UrlEncode($Filter))"
$cnt = 0
Do {
    $people = Invoke-TrelicaRequest -Context $Context -Path $path
    $cnt = $cnt + $people.results.Length
    Write-Host "Processing $cnt..." -ForegroundColor Yellow
    ProcessPeople -Context $Context -People $people.results
    $path = $people.next
} While ($path)
