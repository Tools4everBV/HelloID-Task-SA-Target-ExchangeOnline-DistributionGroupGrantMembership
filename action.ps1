# HelloID-Task-SA-Target-ExchangeOnline-DistributionGroupCreate
###############################################################
# Form mapping
$formObject = @{
    Name               = $form.Name
    DisplayName        = $form.DisplayName
    PrimarySmtpAddress = $form.PrimarySmtpAddress
    Alias              = $form.Alias
}
[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnline action: [DistributionGroupCreate] for: [$($formObject.DisplayName)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop -Verbose:$false -CommandName 'New-DistributionGroup', 'Disconnect-ExchangeOnline'
    $IsConnected = $true

    $createdDistributionGroup = New-DistributionGroup @formObject -ErrorAction Stop

    $auditLog = @{
        Action            = 'CreateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $createdDistributionGroup.ExchangeObjectId
        TargetDisplayName = $createdDistributionGroup.Name
        Message           = "ExchangeOnline action: [DistributionGroupCreate] for: [$($formObject.DisplayName)] executed successfully"
        IsError           = $false
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Information "ExchangeOnline action: [DistributionGroupCreate] for: [$($formObject.DisplayName)] executed successfully"
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'CreateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.DisplayName
        TargetDisplayName = $formObject.DisplayName
        Message           = "Could not execute ExchangeOnline action: [DistributionGroupCreate] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [DistributionGroupCreate] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        $null = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
###############################################################
