# HelloID-Task-SA-Target-ExchangeOnPremises-MailboxUpdateAttributes
###################################################################
# Form mapping
$formObject = @{
    MailboxIdentity  = $form.MailboxIdentity
    DisplayName      = $form.MailboxDisplayName
    CustomAttribute1 = $form.CustomAttribute1
    CustomAttribute2 = $form.CustomAttribute2
    CustomAttribute3 = $form.CustomAttribute3
    CustomAttribute4 = $form.CustomAttribute4
    CustomAttribute5 = $form.CustomAttribute5
    CustomAttribute6 = $form.CustomAttribute6
    CustomAttribute7 = $form.CustomAttribute7
}


[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnPremises action: [MailboxUpdateAttributes] for: [$($formObject.DisplayName)]"

    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'Set-Mailbox'
    $IsConnected = $true

    $SetMailboxParams = @{
        Identity = $formObject.MailboxIdentity
    }
    # Only update the properties that have a value.
    ($formObject.GetEnumerator() | Where-Object { $_.name -ne 'MailboxIdentity' } ).ForEach(  {
            if ($null -ne $_.value) { $SetMailboxParams.Add("$($_.name)", "$($_.value)") }
        })


    $null = Set-Mailbox @SetMailboxParams -ErrorAction Stop

    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.DisplayName
        Message           = "ExchangeOnPremises action: [MailboxUpdateAttributes] for: [$($formObject.DisplayName)] executed successfully"
        IsError           = $false
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Information "ExchangeOnPremises action: [MailboxUpdateAttributes] for: [$($formObject.DisplayName)] executed successfully"
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.DisplayName
        Message           = "Could not execute ExchangeOnPremises action: [MailboxUpdateAttributes] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnPremises action: [MailboxUpdateAttributes] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        Remove-PSSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}

###################################################################
