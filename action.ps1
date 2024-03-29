# HelloID-Task-SA-Target-ExchangeOnline-DistributionGroupGrantMembership
########################################################################
# Form mapping
$formObject = @{
    GroupIdentity   = $form.GroupIdentity
    UsersToAdd = $form.UsersToAdd.Name
}

[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnline action: [DistributionGroupGrantMembership] for: [$($formObject.GroupIdentity)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop -Verbose:$false -CommandName 'Add-DistributionGroupMember', 'Disconnect-ExchangeOnline'
    $IsConnected = $true

    foreach ($user in $formObject.usersToAdd) {
        $AddDistributionGroupMember = @{
            Identity                        = $formObject.GroupIdentity
            Member                          = $user
            BypassSecurityGroupManagerCheck = $true
        }
        try {
            $null = Add-DistributionGroupMember @AddDistributionGroupMember -Confirm:$false -ErrorAction Stop

            $auditLog = @{
                Action            = 'GrantMembership'
                System            = 'ExchangeOnline'
                TargetIdentifier  = $formObject.GroupIdentity
                TargetDisplayName = $formObject.GroupIdentity
                Message           = "ExchangeOnline action: [DistributionGroupGrantMembership] [$($user)] to group [$($formObject.GroupIdentity)] executed successfully"
                IsError           = $false
            }
            Write-Information -Tags 'Audit' -MessageData $auditLog
            Write-Information "ExchangeOnline action: [DistributionGroupGrantMembership] [$($user)] to group [$($formObject.GroupIdentity)] executed successfully"

        } catch {
            if ( $_.Exception.ErrorRecord.CategoryInfo.Reason -eq 'MemberAlreadyExistsException') {
                $auditLog = @{
                    Action            = 'GrantMembership'
                    System            = 'ExchangeOnline'
                    TargetIdentifier  = $formObject.GroupIdentity
                    TargetDisplayName = $formObject.GroupIdentity
                    Message           = "ExchangeOnline action: [DistributionGroupGrantMembership][$($user)] to group [$($formObject.GroupIdentity)] Already Exists"
                    IsError           = $false
                }
                Write-Information -Tags 'Audit' -MessageData $auditLog
                Write-Information "ExchangeOnline action: [DistributionGroupGrantMembership][$($user)] to group [$($formObject.GroupIdentity)] Already Exists"
                Write-Information "Warning message: $($_.Exception.Message)"
                continue
            }
            throw $_
        }
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'GrantMembership'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.Identity
        TargetDisplayName = $formObject.Identity
        Message           = "Could not execute ExchangeOnline action: [DistributionGroupGrantMembership] for: [$($formObject.Identity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [DistributionGroupGrantMembership] for: [$($formObject.Identity)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        $null = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
########################################################################
