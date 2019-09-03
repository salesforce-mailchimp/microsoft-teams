<#
.SYNOPSIS
    Save Azure AD User Credentials

.DESCRIPTION
    Saves a set of Azure AD user credentials.

.OUTPUTS
    verificationAgainstAzureResult
        Verification Against Azure Result
        The result of verifying the credentials against Azure.

    verificationAgainstAzureADResult
        Verification Against Azure AD Result
        The result of verifying the credentials against Azure AD.

    verificationAgainstAzureRMResult
        Verification Against Azure RM Result
        The result of verifying the credentials against Azure RM.

    verificationAgainstExchangeOnlineResult
        Verification Against Exchange Online Result
        The result of verifying the credentials against Exchange Online.

    verificationAgainstMSOnlineResult
        Verification Against MS Online Result
        The result of verifying the credentials against MS Online.

    verificationAgainstOffice365SecurityAndComplianceResult
        Verification Against Office 365 Security and Compliance Center Result
        The result of verifying the credentials against Office 365 Security and Compliance Center.

    verificationResultsMarkdown
        Verification Results Markdown
        A markdown-formatted string containing the verification results.
#>

[CmdletBinding()]
[OutputType()]
param (
    <#
        Credentials Name
        The name used to save the credentials.
    #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$credentialName,

    <#
        Username
        The username.
    #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$username,

    <#
        Password
        The password.
    #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$password,

    <#
        Verify login against Azure?
        Select whether to verify the credentials for login against Azure.
    #>
    [Parameter(Mandatory=$false)]
    [ValidateSet("Yes", "No")]
    [String]$verifyLoginAgainstAzure = "No",

    <#
        Verify login against Azure AD?
        Select whether to verify the credentials for login against Azure AD.
    #>
    [Parameter(Mandatory=$false)]
    [ValidateSet("Yes", "No")]
    [String]$verifyLoginAgainstAzureAD = "No",

    <#
        Verify login against Azure RM?
        Select whether to verify the credentials for login against Azure RM.
    #>
    [Parameter(Mandatory=$false)]
    [ValidateSet("Yes", "No")]
    [String]$verifyLoginAgainstAzureRM = "No",

    <#
        Verify login against Exchange Online?
        Select whether to verify the credentials for login against Exchange Online.
    #>
    [Parameter(Mandatory=$false)]
    [ValidateSet("Yes", "No")]
    [String]$verifyLoginAgainstExchangeOnline = "No",

    <#
        Verify login against MS Online?
        Select whether to verify the credentials for login against MS Online.
    #>
    [Parameter(Mandatory=$false)]
    [ValidateSet("Yes", "No")]
    [String]$verifyLoginAgainstMSOnline = "No",

    <#
        Verify login against Office 365 Security and Compliance Center?
        Select whether to verify the credentials for login against Office 365 Security and Compliance Center.
    #>
    [Parameter(Mandatory=$false)]
    [ValidateSet("Yes", "No")]
    [String]$verifyLoginAgainstOffice365SecurityAndCompliance = "No"
)

####################################################################################################
# Verify the PowerShell edition and version
####################################################################################################

$RequiredPSEdition = "Desktop"
$MinimumPowerShellVersion = [Version]"5.1"
if ($PSVersionTable.PSEdition -ne $RequiredPSEdition) {
    Write-Warning ("The PowerShell edition required is $($RequiredPSEdition). The edition installed is $($PSVersionTable.PSEdition).")
    exit
}
if ($PSVersionTable.PSVersion -lt $MinimumPowerShellVersion) {
    Write-Warning ("The minimum PowerShell version required is $($MinimumPowerShellVersion.ToString()). The current version installed is $($PSVersionTable.PSVersion.ToString()).")
    exit
}

####################################################################################################
# Declarations
####################################################################################################

# Set the default parameter values for the Write-OutputMessage function
$PSDefaultParameterValues["Write-OutputMessage:AppendNewLines"] = 1
$PSDefaultParameterValues["Write-OutputMessage:ToWarning"] = [Switch]::Present

####################################################################################################
# Import functions
####################################################################################################

<#
@(
    "Connect-AzureADAdminAccount",
    "Connect-ExchangeOnlineAdminAccount",
    "Connect-Office365SecurityAndComplianceAdminAccount",
    "Disconnect-ExchangeOnline",
    "Disconnect-Office365SecurityAndCompliance"
) | ForEach-Object -Process {
    . "$($PSScriptRoot)\..\..\functions\$($_).ps1"
}
#>

####################################################################################################
# Initialize the local context
####################################################################################################

$localContextPath = Join-Path $PSScriptRoot "..\local-context.ps1"
if (Test-Path $localContextPath) {
    . $localContextPath
}

####################################################################################################
# The main program
####################################################################################################

try {
    # Verify the logins
    # TODO: Refactor out as this is shared with Verify Azure AD User Credentials
    $allLoginsSucceeded = $true
    $azureLoginSucceeded = $null
    $azureADLoginSucceeded = $null
    $azureRMLoginSucceeded = $null
    $exchangeOnlineLoginSucceeded = $null
    $msOnlineLoginSucceeded = $null
    $office365SecurityAndComplianceLoginSucceeded = $true
    <#
    if ($verifyLoginAgainstAzure -eq "Yes") {
        try {
            Import-Module Az
            Connect-AzAccount -Credential ([PSCredential]::new(
                $username,
                ($password | ConvertTo-SecureString -AsPlainText -Force)
            )) -ErrorAction Stop | Out-Null
            Write-Information "Login against Azure verified successfully."
            $azureLoginSucceeded = $true
        }
        catch {
            Write-Warning "Exception occurred while verifying the login against Azure.`r`n$($_.Exception.Message)"
            $azureLoginSucceeded = $false
            $allLoginsSucceeded = $false
        }
        finally {
            Disconnect-AzAccount
            Remove-Module Az*
        }
    }
    if ($verifyLoginAgainstAzureAD -eq "Yes") {
        try {
            Import-Module AzureAD
            Connect-AzureADAdminAccount -Username $username -Password ($password | ConvertTo-SecureString -AsPlainText -Force) -ErrorAction Stop | Out-Null
            Write-Information "Login against Azure AD verified successfully."
            $azureADLoginSucceeded = $true
        }
        catch {
            Write-Warning "Exception occurred while verifying the login against Azure AD.`r`n$($_.Exception.Message)"
            $azureADLoginSucceeded = $false
            $allLoginsSucceeded = $false
        }
    }
    if ($verifyLoginAgainstAzureRM -eq "Yes") {
        try {
            Import-Module AzureRM
            Connect-AzureRMAccount -Credential ([PSCredential]::new(
                $username,
                ($password | ConvertTo-SecureString -AsPlainText -Force)
            )) -Environment "AzureCloud" -ErrorAction Stop | Out-Null
            Write-Information "Login against Azure RM verified successfully."
            $azureRMLoginSucceeded = $true
        }
        catch {
            Write-Warning "Exception occurred while verifying the login against Azure RM.`r`n$($_.Exception.Message)"
            $azureRMLoginSucceeded = $false
            $allLoginsSucceeded = $false
        }
    }
    if ($verifyLoginAgainstExchangeOnline -eq "Yes") {
        try {
            Connect-ExchangeOnlineAdminAccount -Username $username -Password ($password | ConvertTo-SecureString -AsPlainText -Force) -ErrorAction Stop | Out-Null
            Write-Information "Login against Exchange Online verified successfully."
            $exchangeOnlineLoginSucceeded = $true
        }
        catch {
            Write-Warning "Exception occurred while verifying the login against Exchange Online.`r`n$($_.Exception.Message)"
            $exchangeOnlineLoginSucceeded = $false
            $allLoginsSucceeded = $false
        }
        finally {
            Disconnect-ExchangeOnline
        }
    }
    if ($verifyLoginAgainstMSOnline -eq "Yes") {
        try {
            Connect-MsolService -Credential ([PSCredential]::new(
                $username,
                ($password | ConvertTo-SecureString -AsPlainText -Force)
            )) -ErrorAction Stop | Out-Null
            Write-Information "Login against MS Online verified successfully."
            $msOnlineLoginSucceeded = $true
        }
        catch {
            Write-Warning "Exception occurred while verifying the login against MS Online.`r`n$($_.Exception.Message)"
            $msOnlineLoginSucceeded = $false
            $allLoginsSucceeded = $false
        }
    }
    if ($verifyLoginAgainstOffice365SecurityAndCompliance -eq "Yes") {
        try {
            Connect-Office365SecurityAndComplianceAdminAccount -Username $username -Password ($password | ConvertTo-SecureString -AsPlainText -Force) -ErrorAction Stop | Out-Null
            Write-Information "Login against Office 365 Security and Compliance Center verified successfully."
            $office365SecurityAndComplianceLoginSucceeded = $true
        }
        catch {
            Write-Warning "Exception occurred while verifying the login against Office 365 Security and Compliance Center.`r`n$($_.Exception.Message)"
            $office365SecurityAndComplianceLoginSucceeded = $false
            $allLoginsSucceeded = $false
        }
        finally {
            Disconnect-Office365SecurityAndCompliance
        }
    }

    # Output a summary of the verifications
    # TODO: Refactor out as this is shared with Verify Azure AD User Credentials
    Write-Information "##### Verification Results #####"
    $context.outputs.verificationResultsMarkdown = ""
    if ($null -eq $azureLoginSucceeded) {
        $context.outputs.verificationAgainstAzureResult = "Not run"
        $context.outputs.verificationResultsMarkdown += "`n`Azure: Not run"
        Write-Information "Azure: Not run"
    }
    elseif ($azureLoginSucceeded) {
        $context.outputs.verificationAgainstAzureResult = "Succeeded"
        $context.outputs.verificationResultsMarkdown += "`n`Azure: Succeeded"
        Write-Information "Azure: Succeeded"
    }
    else {
        $context.outputs.verificationAgainstAzureResult = "Failed"
        $context.outputs.verificationResultsMarkdown += "`n`nAzure: Failed"
        Write-Information "Azure: Failed"
    }
    if ($null -eq $azureADLoginSucceeded) {
        $context.outputs.verificationAgainstAzureADResult = "Not run"
        $context.outputs.verificationResultsMarkdown += "`n`nAzure AD: Not run"
        Write-Information "Azure AD: Not run"
    }
    elseif ($azureADLoginSucceeded) {
        $context.outputs.verificationAgainstAzureADResult = "Succeeded"
        $context.outputs.verificationResultsMarkdown += "`n`nAzure AD: Succeeded"
        Write-Information "Azure AD: Succeeded"
    }
    else {
        $context.outputs.verificationAgainstAzureADResult = "Failed"
        $context.outputs.verificationResultsMarkdown += "`n`nAzure AD: Failed"
        Write-Information "Azure AD: Failed"
    }
    if ($null -eq $azureRMLoginSucceeded) {
        $context.outputs.verificationAgainstAzureRMResult = "Not run"
        $context.outputs.verificationResultsMarkdown += "`n`nAzure RM: Not run"
        Write-Information "Azure RM: Not run"
    }
    elseif ($azureRMLoginSucceeded) {
        $context.outputs.verificationAgainstAzureRMResult = "Succeeded"
        $context.outputs.verificationResultsMarkdown += "`n`nAzure RM: Succeeded"
        Write-Information "Azure RM: Succeeded"
    }
    else {
        $context.outputs.verificationAgainstAzureRMResult = "Failed"
        $context.outputs.verificationResultsMarkdown += "`n`nAzure RM: Failed"
        Write-Information "Azure RM: Failed"
    }
    if ($null -eq $exchangeOnlineLoginSucceeded) {
        $context.outputs.verificationAgainstExchangeOnlineResult = "Not run"
        $context.outputs.verificationResultsMarkdown += "`n`nExchange Online: Not run"
        Write-Information "Exchange Online: Not run"
    }
    elseif ($exchangeOnlineLoginSucceeded) {
        $context.outputs.verificationAgainstExchangeOnlineResult = "Succeeded"
        $context.outputs.verificationResultsMarkdown += "`n`nExchange Online: Succeeded"
        Write-Information "Exchange Online: Succeeded"
    }
    else {
        $context.outputs.verificationAgainstExchangeOnlineResult = "Failed"
        $context.outputs.verificationResultsMarkdown += "`n`nExchange Online: Failed"
        Write-Information "Exchange Online: Failed"
    }
    if ($null -eq $msOnlineLoginSucceeded) {
        $context.outputs.verificationAgainstMSOnlineResult = "Not run"
        $context.outputs.verificationResultsMarkdown += "`n`nMS Online: Not run"
        Write-Information "MS Online: Not run"
    }
    elseif ($msOnlineLoginSucceeded) {
        $context.outputs.verificationAgainstMSOnlineResult = "Succeeded"
        $context.outputs.verificationResultsMarkdown += "`n`nMS Online: Succeeded"
        Write-Information "MS Online: Succeeded"
    }
    else {
        $context.outputs.verificationAgainstMSOnlineResult = "Failed"
        $context.outputs.verificationResultsMarkdown += "`n`nMS Online: Failed"
        Write-Information "MS Online: Failed"
    }
    if ($null -eq $office365SecurityAndComplianceLoginSucceeded) {
        $context.outputs.verificationAgainstOffice365SecurityAndComplianceResult = "Not run"
        $context.outputs.verificationResultsMarkdown += "`n`nOffice 365 Security and Compliance Center: Not run"
        Write-Information "Office 365 Security and Compliance Center: Not run"
    }
    elseif ($office365SecurityAndComplianceLoginSucceeded) {
        $context.outputs.verificationAgainstOffice365SecurityAndComplianceResult = "Succeeded"
        $context.outputs.verificationResultsMarkdown += "`n`nOffice 365 Security and Compliance Center: Succeeded"
        Write-Information "Office 365 Security and Compliance Center: Succeeded"
    }
    else {
        $context.outputs.verificationAgainstOffice365SecurityAndComplianceResult = "Failed"
        $context.outputs.verificationResultsMarkdown += "`n`nOffice 365 Security and Compliance Center: Failed"
        Write-Information "Office 365 Security and Compliance Center: Failed"
    }
    #>

    # Save the credentials if all of the logins were successful
    if ($allLoginsSucceeded) {
        # Generate the json file contents containing the credentials
        $credentialsJson = @{
            username = $username
            password = $password
        } | ConvertTo-Json

        # Generate the credentials file name
        $credentialsFileName = "AzureADUserCredentials_" + $credentialName + ".json"

        # Save the package text
        $context.SaveVendorText($credentialsFileName, $credentialsJson)
        $context.outputs.verificationResultsMarkdown += "`n`nThe provided credentials have been saved as '$($credentialName)'."
        Write-Information "The provided credentials have been saved as '$($credentialName)'."
    }

    # Output a message explaining the outcome
    else {
        $context.outputs.verificationResultsMarkdown += "`n`nThe provided credentials have not been verified successfully. The provided credentials have not been saved."
        Write-Warning "The provided credentials have not been verified successfully. The provided credentials have not been saved."
    }
}
catch {
    Write-Warning "Exception occurred on line $($_.InvocationInfo.ScriptLineNumber):`r`n$($_.Exception.Message)"
}
finally {
    # Nothing to clean up
}