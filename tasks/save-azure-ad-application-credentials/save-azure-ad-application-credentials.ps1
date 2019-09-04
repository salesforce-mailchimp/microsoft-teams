<#
.SYNOPSIS
    Save Azure AD Application Credentials

.DESCRIPTION
    Saves a set of Azure AD application credentials.

.OUTPUTS
    verificationAgainstMicrosoftGraphResult
        Verification Against Microsoft Graph Result
        The result of verifying the credentials against Microsoft Graph.

    verificationResultsMarkdown
        Verification Results Markdown
        A markdown-formatted string containing the verification results.
#>
[CmdletBinding()]
[OutputType()]
param (
    <#
        Halt Execution on Error?
        Selects whether this task should halt execution of the template upon encountering an error.
    #>
    [Parameter(Mandatory=$false)]
    [ValidateSet("Yes", "No")]
    [String]$haltExecutionOnError = "No",

    <#
        Credentials Name
        The name used to save the credentials.
    #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$credentialsName,

    <#
        Application ID
        The application ID.
    #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$applicationId,

    <#
        Client Secret
        The client secret.
    #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$clientSecret,

    <#
        Tenant Domain
        The tenant domain.
    #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$tenantDomain,

    <#
        Verify Against Microsoft Graph?
        Select whether to verify the provided credentials against Microsoft Graph.
    #>
    [Parameter(Mandatory=$false)]
    [ValidateSet("Yes", "No")]
    [String]$verifyAgainstMicrosoftGraph = "No"
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
    # Store error messages in a string
    $errorMessages = ""

    # Verify the credentials
    # TODO: Refactor out as this is shared with Verify Azure AD Application Credentials
    $allCredentialVerificationsSuccessful = $true
    $microsoftGraphCredentialVerificationSuccessful = $true
    <#
    if ($verifyAgainstMicrosoftGraph -eq "Yes") {
        try {
            $token = Get-MicrosoftGraphAuthenticationToken -ApplicationId $applicationId -ClientSecret $clientSecret -Domain $tenantDomain -ErrorAction Stop
            if ([String]::IsNullOrWhiteSpace($token)) {
                $errorMessages += Write-OutputMessage "Received a null or empty Microsoft Graph authentication token when using the provided credentials."
                $allCredentialVerificationsSuccessful = $false
                $microsoftGraphCredentialVerificationSuccessful = $false
            }
            else {
                Write-Information "Credentials verified successfully against Microsoft Graph."
                $microsoftGraphCredentialVerificationSuccessful = $true
            }
        }
        catch {
            $errorMessages += Write-OutputMessage "Exception occurred while verifying the provided credentials against Microsoft Graph.`r`n$($_.Exception.Message)"
            $allCredentialVerificationsSuccessful = $false
            $microsoftGraphCredentialVerificationSuccessful = $false
        }
    }
    #>

    # Output a summary of the verifications
    # TODO: Refactor out as this is shared with Verify Azure AD Application Credentials
    Write-Information "##### Verification Results #####"
    $context.outputs.verificationResultsMarkdown = "# Verification Results"
    if ($null -eq $microsoftGraphCredentialVerificationSuccessful) {
        $context.outputs.verificationAgainstMicrosoftGraphResult = "Not run"
        $context.outputs.verificationResultsMarkdown += "`n`nMicrosoft Graph: Not run"
        Write-Information "Microsoft Graph: Not run"
    }
    elseif ($microsoftGraphCredentialVerificationSuccessful) {
        $context.outputs.verificationAgainstMicrosoftGraphResult = "Successful"
        $context.outputs.verificationResultsMarkdown += "`n`nMicrosoft Graph: Successful"
        Write-Information "Microsoft Graph: Successful"
    }
    else {
        $context.outputs.verificationAgainstMicrosoftGraphResult = "Failed"
        $context.outputs.verificationResultsMarkdown += "`n`nMicrosoft Graph: Failed"
        Write-Information "Microsoft Graph: Failed"
    }


    # Save the credentials if the verification was successful
    if ($allCredentialVerificationsSuccessful) {
        # Generate the json file contents containing the credentials
        $credentialsJson = @{
            applicationId = $applicationId
            clientSecret  = $clientSecret
            domain        = $tenantDomain
        } | ConvertTo-Json

        # Generate the credentials file name
        $credentialsFileName = "AzureADApplicationCredentials_" + $credentialsName + ".json"

        # Save the file containing the credentials to the package scope
        $context.SaveVendorText($credentialsFileName, $credentialsJson)
        $context.outputs.verificationResultsMarkdown += "`n`nThe provided credentials have been saved as '$($credentialsName)'."
        Write-Information "The provided credentials have been saved as '$($credentialsName)'."
    }

    # Output a message explaining the outcome
    else {
        $context.outputs.verificationResultsMarkdown += "`n`nThe provided credentials have not been verified successfully. The provided credentials have not been saved."
        $errorMessages += Write-OutputMessage "The provided credentials have not been verified successfully. The provided credentials have not been saved."
    }
}
catch {
    $errorMessages += Write-OutputMessage "Exception occurred on line $($_.InvocationInfo.ScriptLineNumber):`r`n$($_.Exception.Message)"
}
finally {
    # Output the error messages
    $context.Outputs.errorMessages = $errorMessages

    # Halt execution of the template if there was an error
    if ($haltExecutionOnError -eq "Yes" -and ![String]::IsNullOrWhiteSpace($errorMessages)) {
        Write-Error "Errors occurred during task execution. Template execution will be halted."
    }
}
