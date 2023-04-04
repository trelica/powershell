[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)][PSTypeName('Trelica.Context')]$Context
)
<# 
    .SYNOPSIS
    Configure this script to call Trelica on a schedule and then process waiting
    workflow runs to create or disable AD user accounts.
    .INPUTS
    An authenticated Trelica Powershell conext object, created by Initialize-TrelicaCredentials
    .EXAMPLE
    Run this passing in a Trelica Context:
    Initialize-TrelicaCredentials | Get-TrelicaContext | .\ADWorkflowSync.ps1
    .LINK
    Read more online: https://help.trelica.com/hc/en-us/articles/10215817603229
#>

# stop on first error
$ErrorActionPreference = "Stop"

# Set to $true to do a dry run, not connecting to AD or posting results to Trelica
$DRY_RUN = $true

# Set to $true to write out JSON received from Trelica 
$DEBUG = $false

# Core settings
$ADHOST = "XXXXXXXXXXX"
$PROVISIONING_WORKFLOW_LIST = @("New Hire") 
$DEPROVISIONING_WORKFLOW_LIST = @("Terminate Users")
$STEP_NAME = "Waiting for Active Directory Synch"
$ACTION_NAME = "Continue"

# Specific settings for provisioning workflows
$DOMAIN = "corp.example.org"
$DEFAULT_OU = "Users/All"
$TEAM_TO_OU_MAP = @{
    "Engineering"="Users/Dev"
    "DevOps"="Users/Dev"
}
$NEW_USER_GROUPS = @("AllUsers")



<# --------------------------------------- #>
Function Get-DN() {
    [OutputType([String])]
    Param(
        [Parameter(Mandatory = $true)][System.Object]$Domain,
        [Parameter(Mandatory = $true)][System.Object]$OUPath
    )

    $dn = ""
    $ouParts = $OUPath.Split("/")
    for ($i = 0; $i -lt $ouParts.length; $i++) {
        if ("" -ne $ouParts[$i]) {
            $dn += ",OU=$($ouParts[$i].Trim())"
        }
    }
    $dn = $dn.TrimStart(",")
    foreach ($p in $Domain.Split(".")) {
        $dn += ",dc=$p"
    }
    $dn 
}

Function ProcessProvisioningWorkflowRun() {
    Param(
        [Parameter(Mandatory = $true)][PSTypeName('Trelica.Context')]$Context,
        [Parameter(Mandatory = $true)][System.Object]$Run,
        [Parameter(Mandatory = $true)][System.Object]$ActionName,
        [Parameter(Mandatory = $true)][System.Object]$ActionStepId
    )
    $createUserStep = $null
    $email = $null

    if ($DEBUG) {
        Write-Host ($Run | ConvertTo-Json -Depth 10)
    }
    
    # find the step we're waiting on (probably it's the last one)
    for ($i = ($Run.steps.length-1); $i -ge 0; $i--) {
        $step = $Run.steps[$i]
        if ($step.id -eq $ActionStepId -and $step.status -eq "Waiting") {
            $createUserStep = $step
            continue
        }
        # we then go back one further to get the prior context as maybe the email got
        # set during an earlier Create user workflow step
        if ($null -ne $createUserStep) {
            $person = $step.completed.newContext.person
            if ($null -ne $person) {
                $email = $person.email
                Write-Host "INFO: email taken from context = $($person.email) [$($person.name)]"
                break
            }
        }
    }
    if (!$createUserStep) {
        Write-Information "ASSERTION FAILURE: cannot find action step id $ActionStepId"
        return $false
    }
    
    $person = $Run.context.person
    if (!$email) {
        # take the email from the context
        $email = $person.email
    }

    # username is assumed to be the part before the @ character
    $userName = ($email) -replace '(.*)@.*', '$1'
    # if the SamAccountName is going to be beyond 20 chars then switch
    # to (initial)(surname), and then truncate if even that's too long
    if (20 -le $userName.length) {
        $userName = ($email) -replace '(.)(.*)\.(.*)@.*', '$1$3'
        if (20 -le $userName.length) {
            $userName = $userName.substring(0,20)
        }
    }
    $userPrincipal = $userName + "@" + $DOMAIN

    # figure out the DN string for the OU, based on the Trelica team
    $team = $person.team
    if ($null -eq $team) {
        Write-Host "WARNING: No team sent - defaulting to '$DEFAULT_OU' in order to assign an OU"
        $path = $DEFAULT_OU
    } elseif (-not $TEAM_TO_OU_MAP.ContainsKey($team)) {
        Write-Host "WARNING: Team '$team' not found in map - defaulting to '$DEFAULT_OU' in order to assign an OU"
        $path = $DEFAULT_OU
    } else {
        $path = Get-DN -OUPath $TEAM_TO_OU_MAP[$team] -Domain $domain
    }

    # get the password from the workflow run variables context
    $password = $Run.context.variables.password

    Write-Host "INFO: Creating AD user '$userName' ($userPrincipal) under '$path' with email '$email'..."
    if ([string]::IsNullOrWhiteSpace($password)) {
        Write-Host "WARNING: no passsword set"
    }

    if (-not $DRY_RUN) {        
        Write-Host "-- checking if user exists..."
        $adUser = Get-ADUser -LDAPFilter "(sAMAccountName=$userName)"
        if ($null -ne $adUser) {
            Write-Host "'${userName}' already exists in AD"
        } else {
            # ACTION: Create user 
            Write-Host "-- creating new user..."
            try
            {
                New-ADUser -Name $person.name -GivenName $person.firstName -Surname $person.lastName `
                    -SamAccountName $userName `
                    -UserPrincipalName $userPrincipal `
                    -EmailAddress $email `
                    -Description "Trelica Workflow Run ID $($Run.id)" `
                    -Path $path `
                    -Enabled $true `
                    -ChangePasswordAtLogon $false `
                    -AccountPassword (ConvertTo-SecureString "$password" -AsPlainText -force) `
                    -passThru `
                    -DisplayName $person.name `
                    -ErrorAction Stop
            }
            catch 
            {
                Write-Host "ERROR: An error occurred creating the AD user - aborting:"
                Write-Host $_
                return
            }

            # ACTION: Add to groups
            Write-Host "-- adding to groups..."
            
            forEach ($group In $NEW_USER_GROUPS) {
                try 
                {
                    Write-Host "-- adding user '$userName' to group '$group'..."
                    Add-ADGroupMember -Identity $group -Members $userName -Confirm:$False -ErrorAction Stop
                }
                catch
                {
                    Write-Host "ERROR: An error occurred adding the user to the group:"
                    Write-Host $_
                    continue
                }
            }
        }
    }

    # now return the URI of the action we want
    return ($createUserStep.waiting.actions | Where-Object { $_.name -eq $ActionName }).href
}

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
    
    # we hope this person really exists in AD
    $userName = ($person.email) -replace '(.*)@.*', '$1'
    # $auditDescription = "Trelica Workflow Run ID $($Run.id)" 

    Write-Host "INFO: Deprovisioning AD user '$userName'..."
    
    if (-not $DRY_RUN) {
        try {
            $adUser = Get-ADUser -Identity $userName -ErrorAction Stop -Properties * 
        }
        catch {
            Write-Host "ERROR: An error occurred reading the AD user - aborting:"
            Write-Host $_
            return $false
        }

        # ACTION: Remove from all groups
        try {
            $groups = (Get-ADUser -Identity $userName -Properties memberOf -ErrorAction Stop).memberOf
        }
        catch {
            Write-Host "ERROR: An error occurred getting groups for the AD user - aborting:"
            Write-Host $_
            return $false
        }
        forEach ($group In $groups) {
            Write-Host "-- removing user '$userName' from group '$group'..."
            try {
                Remove-ADGroupMember -Identity $group -Members $userName -Confirm:$FALSE -ErrorAction Stop                
            }
            catch {
                Write-Host "ERROR: An error occurred removing the user from the group - aborting:"
                Write-Host $_
                return $false
            }
        }

        # ACTION: Disable the account
        Write-Host "-- disabling user '$userName'..."
        try {
            Disable-ADAccount -Identity $userName -ErrorAction Stop           
        }
        catch {
            Write-Host "ERROR: An error occurred disabling the user - aborting:"
            Write-Host $_
            return $false    
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
        } else {
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
    $actionStepId = (($workflow.steps | Where-Object {$_.name -eq $STEP_NAME}).id)
    if ([string]::IsNullOrWhiteSpace($actionStepId)) {
        throw "Step not found '$STEP_NAME'"
    }

    # pull back any runs that are waiting
    Write-Host "Requesting runs for Workflow '$WorkflowName' (ID $id) and Action Step ID $actionStepId..." 
    $filter = @{
        filter = 'status ne "Terminated" and steps[status eq "Waiting" and id eq "' + $actionStepId + '"]'
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
    Write-Host "********************************************************"
    Write-Host "DRY RUN: NOT CONNECTING TO AD OR POSTING BACK TO TRELICA"
    Write-Host "********************************************************"
} else { 
    Write-Host "Connecting to $ADHOST..."
    try
    {
        $S = New-PSSession -ComputerName $ADHOST
    }
    catch 
    {
        Write-Host "An error occurred connecting to AD - aborting:"
        Write-Host $_
        break
    }
    Import-Module -PSsession $S -Name ActiveDirectory
}

Write-Host
Write-Host "DEPROVISIONING WORKFLOWS"
Write-Host "========================"
foreach ($workflowName in $DEPROVISIONING_WORKFLOW_LIST) {
    ProcessWorkflow -Context $Context -Name $workflowName -ProcessWith ProcessDeprovisioningWorkflowRun
}

Write-Host 
Write-Host "PROVISIONING WORKFLOWS"
Write-Host "======================"
foreach ($workflowName in $PROVISIONING_WORKFLOW_LIST) {
    ProcessWorkflow -Context $Context -Name $workflowName -ProcessWith ProcessProvisioningWorkflowRun
}

Write-Host
Write-Host "****"
Write-Host "DONE"
Write-Host "****"