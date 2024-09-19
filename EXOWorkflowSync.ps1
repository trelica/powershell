[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)][PSTypeName('Trelica.Context')]$Context,
    [Parameter(Mandatory = $true)][securestring]$CertificatePassword
)
<# 
    .SYNOPSIS
    Configure this script to call Trelica on a schedule and then process waiting
    workflow runs to manage Exchange Online.
    .INPUTS
    An authenticated Trelica Powershell conext object, created by Initialize-TrelicaCredentials
    .EXAMPLE
    Run this passing in a Trelica Context:
    Initialize-TrelicaCredentials | Get-TrelicaContext | ./EXOWorkflowSync.ps1
    .LINK
    Read more online: https://help.trelica.com/hc/en-us/articles/17910775949469
#>

# stop on first error
$ErrorActionPreference = "Stop"

# Set to $true to do a dry run, not connecting to EXO or posting results to Trelica
$DRY_RUN = $false

# Set to $true to write out JSON received from Trelica 
$DEBUG = $false

# Core settings
$CERT_PATH = "/Users/xxxxx/scripts/mycert.pfx"
$ORG = "xxxxx.onmicrosoft.com"
$APPID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$DEPROVISIONING_WORKFLOW_LIST = @("Terminate Users")
$STEP_NAME = "Signal"
$ACTION_NAME = "Continue"

<# --------------------------------------- #>

Function ProcessDeprovisioningWorkflowRun() {
    Param(
        [Parameter(Mandatory = $true)][PSTypeName('Trelica.Context')]$Context,
        [Parameter(Mandatory = $true)][System.Object]$Run,
        [Parameter(Mandatory = $true)][System.Object]$ActionName,
        [Parameter(Mandatory = $true)][System.Object]$ActionStepId
    )

    $actionStep = $Run.steps | Where-Object { $_.id -eq "$ActionStepId" -and $_.status -eq "Waiting" } 
    if (!$actionStep) {
        Write-Information "ASSERTION FAILURE: cannot find action step id $ActionStepId"
        return $false
    }

    # Go off and deprovision in Active Directory
    $person = $Run.context.person
    if ($DEBUG) {
        Write-Host ($person | ConvertTo-Json -Depth 10)
    }

    $email = $person.email
        
    if (-not $DRY_RUN) {
        try {
            Set-Mailbox $email -Type Shared
        }
        catch {
            Write-Host "ERROR: An error marking the mailbox as shared - Aborting: $_"
            return $false
        }

        # Declare the distributionLists variable outside of the try-catch block
        $distributionLists = @()

        try {
            # Attempt to get the distribution lists
            $distributionLists = Get-DistributionList
        }
        catch {
            # Handle error if fetching distribution lists fails
            Write-Host "Failed to retrieve distribution lists. Error: $_"
            # Optionally exit the script if this is a critical failure
            return
        }

        # If the distribution lists were fetched, iterate over each and try removing the group
        foreach ($group in $distributionLists) {
            try {
                # Attempt to remove the group
                Remove-DistributionGroup -Identity $group.Identity
                Write-Host "Successfully removed group: $($group.Name)"
            }
            catch {
                # Handle error and include group name in the error message
                Write-Host "Failed to remove group: $($group.Name). Error: $_"
            }
        }
    }

    # now return the URI of the action we want
    return ($actionStep.waiting.actions | Where-Object { $_.name -eq $ActionName }).href
}

Function ProcessWorkflowRuns() {
    Param(
        [Parameter(Mandatory = $true)][PSTypeName('Trelica.Context')]$Context,
        [Parameter(Mandatory = $true)][System.Object]$Runs,
        [Parameter(Mandatory = $true)][System.Object]$ActionName,
        [Parameter(Mandatory = $true)][System.Object]$ActionStepId,
        [Parameter(Mandatory = $true)][System.Object]$ProcessWith
    )
    Foreach ($run in $Runs) {
        $now = (Get-Date -Format "yyyy-mm-dd HH:mm K")
        $msg = "> Processing run ID $($run.id): '$($run.name)' at $now"
        Write-Host $msg 
        Write-Host $("-" * $msg.length)
        $successUri = & $ProcessWith -Context $Context -Run $run -ActionName $ActionName -ActionStepId $ActionStepId
        
        if ($false -eq $successUri) {
            # failed in some way, but try the next run anyway
            continue
        }

        if ($null -eq $successUri) {
            Write-Host ($actionStep | ConvertTo-Json -Depth 4)
            Write-Host "WARNING: cannot update Trelica as '$ActionName' action not found in workflow signal step for run ID $($Run.id): '$($Run.name)'"
            continue
        }
    
        # report back to Trelica that we're done
        Write-Host "POSTing '$ActionName' action for run ID $($Run.id): '$($Run.name)'..."
        if ($DRY_RUN) {
            Write-Host "^^^^ DRY RUN: NOT EXECUTING"
        }
        else {
            Invoke-TrelicaRequest -Context $Context -Path $successUri -Method POST
        }
    }
}


Function Get-WorkflowFromName() {
    Param(
        [Parameter(Mandatory = $true)][PSTypeName('Trelica.Context')]$Context,
        [Parameter(Mandatory = $true)][System.Object]$Name
    )
    $workflows = Invoke-TrelicaRequest -Context $Context -Path '/api/workflows/v1' -QueryStringParams @{ filter = 'name eq "' + $Name + '"' }
    if ($workflows.results.length -eq 0) {
        throw "Workflow '$Name' not found"
    }
    if ($workflows.results.length -gt 1) {
        throw "Multiple workflows returned for '$Name'"
    }
    return $workflows.results
}


# get workflow
Function ProcessWorkflow() {
    Param(
        [Parameter(Mandatory = $true)][PSTypeName('Trelica.Context')]$Context,
        [Parameter(Mandatory = $true)][System.Object]$Name,
        [Parameter(Mandatory = $true)][System.Object]$ProcessWith
    )
    $workflow = (Get-WorkflowFromName -Context $Context -Name $Name)

    # get ID of a named workflow and step
    $id = $workflow.id
    $actionStepId = (($workflow.steps | Where-Object { $_.name -eq $STEP_NAME }).id)
    if ([string]::IsNullOrWhiteSpace($actionStepId)) {
        throw "Step not found '$STEP_NAME'"
    }

    # pull back any runs that are waiting
    Write-Host "Requesting runs for Workflow '$WorkflowName' (ID $id) and Action Step ID $actionStepId..." 
    $filter = @{
        filter    = 'status ne "Terminated" and steps[status eq "Waiting" and id eq "' + $actionStepId + '"]'
        variables = 'password' 
    }
    $runs = Invoke-TrelicaRequest -Context $Context -Path "/api/workflows/v1/$id/runs?limit=50" -QueryStringParams $filter
    while ($runs.next) {
        ProcessWorkflowRuns -Context $Context -Runs $runs.results -ActionName $ACTION_NAME -ActionStepId $actionStepId -ProcessWith $ProcessWith
        $runs = Invoke-TrelicaRequest -Context $Context -Path $runs.next
    }
    # last page
    ProcessWorkflowRuns -Context $Context -Runs $runs.results -ActionName $ACTION_NAME -ActionStepId $actionStepId -ProcessWith $ProcessWith
}


<# --------------------------------------- #>

# Connect and then process provisioning and deprovisioning workflows

if ($DRY_RUN) {
    Write-Host "*********************************************************"
    Write-Host "DRY RUN: NOT CONNECTING TO EXO OR POSTING BACK TO TRELICA"
    Write-Host "*********************************************************"
}
else { 
    Write-Host "Connecting to $ORG..."
    try {
        Connect-ExchangeOnline -CertificateFilePath "$CERT_PATH" -CertificatePassword $CertificatePassword -AppID $APPID -Organization $ORG
    }
    catch {
        Write-Host "An error occurred connecting to EXO - aborting:"
        Write-Host $_
        break
    }
}

Write-Host
Write-Host "DEPROVISIONING WORKFLOWS"
Write-Host "========================"
foreach ($workflowName in $DEPROVISIONING_WORKFLOW_LIST) {
    ProcessWorkflow -Context $Context -Name $workflowName -ProcessWith ProcessDeprovisioningWorkflowRun
}

Disconnect-ExchangeOnline -Confirm:$false

Write-Host
Write-Host "****"
Write-Host "DONE"
Write-Host "****"