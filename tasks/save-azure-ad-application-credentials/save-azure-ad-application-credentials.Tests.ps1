Describe "voleer/office-365/save-azure-ad-application-credentials" -Tag "integration" {

    # Declare the task script path
    $taskScriptPath = "$($PSScriptRoot)\save-azure-ad-application-credentials.ps1"

    # Import the Voleer local context
    . "$($PSScriptRoot)\..\local-context.ps1"

    Context "when there are no issues" {

        It "verifies and saves the credentials" {
            # Verify that the credentials have not been saved yet
            $credentialsFileName = "AzureADApplicationCredentials_SaveAzureADApplicationCredentialsTest.json"
            $context.SaveVendorText($credentialsFileName, "")
            $context.GetVendorText($credentialsFileName) | Should BeNullOrEmpty

            # Call the task script
            Invoke-Expression ". $($taskScriptPath) -credentialsName 'SaveAzureADApplicationCredentialsTest' -applicationId '$($env:MicrosoftGraphApplicationID)' -clientSecret '$($env:MicrosoftGraphClientSecret)' -tenantDomain '$($env:MicrosoftGraphTenantDomain)'" -WarningVariable warningVariable

            # Verify that the credentials have been saved
            $credentialObject = $context.GetVendorText($credentialsFileName) | ConvertFrom-Json
            $credentialObject.applicationId | Should Be $env:MicrosoftGraphApplicationID
            $credentialObject.clientSecret | Should Be $env:MicrosoftGraphClientSecret
            $credentialObject.domain | Should Be $env:MicrosoftGraphTenantDomain

            # Verify the output
            $warningVariable | Should BeNullOrEmpty

            # Remove the saved credentials
            $context.SaveVendorText($credentialsFileName, "")
        }
    }
}

Describe "voleer/office-365/save-azure-ad-application-credentials" -Tag "task", "unit" {

    # Declare the task script path
    $taskScriptPath = "$($PSScriptRoot)\save-azure-ad-application-credentials.ps1"

    # Import the Voleer local context
    . "$($PSScriptRoot)\..\local-context.ps1"

    # Import the required functions
    . "$($PSScriptRoot)\..\..\functions\Write-OutputMessage.ps1"

    # Declare functions and mocks
    function Import-VoleerPackageFunction {
        param ($FunctionName)
    }
    function Get-MicrosoftGraphAuthenticationToken {
        param ($ApplicationID, $ClientSecret, $TenantDomain)
    }

    Context "when there are no issues" {
        # Declare mocks
        Mock Get-MicrosoftGraphAuthenticationToken {
            return "token"
        }

        It "saves the credentials" {
            # Declare the task inputs
            $taskInputs = @{
                applicationId               = "Application ID"
                clientSecret                = "Client Secret"
                tenantDomain                = "Tenant Domain"
                verifyAgainstMicrosoftGraph = "Yes"
            }

            # Verify that the credentials are not present
            $credentialsFileName = "AzureADApplicationCredentials_SaveAzureADApplicationCredentialsTest.json"
            $context.SaveVendorText($credentialsFileName, "")
            $context.GetVendorText($credentialsFileName) | Should BeNullOrEmpty

            # Call the task script
            Invoke-Expression ". $($taskScriptPath) -credentialsName 'SaveAzureADApplicationCredentialsTest' @taskInputs" -WarningVariable warningVariable

            # Verify the mocks
            Assert-MockCalled Get-MicrosoftGraphAuthenticationToken -Times 1 -Exactly -ParameterFilter {
                $ApplicationID -eq "Application ID" -and
                $ClientSecret -eq "Client Secret" -and
                $TenantDomain -eq "Tenant Domain"
            } -Scope It

            # Verify the outputs
            $warningVariable | Should BeNullOrEmpty
            $context.Outputs.errorMessages | Should BeNullOrEmpty
            $context.Outputs.verificationAgainstMicrosoftGraphResult | Should Be "Successful"

            # Verify the credentials
            $credentialsObject = $context.GetVendorText($credentialsFileName) | ConvertFrom-Json
            $credentialsObject.applicationId | Should Be "Application ID"
            $credentialsObject.clientSecret | Should Be "Client Secret"
            $credentialsObject.domain | Should Be "Tenant Domain"

            # Remove the credentials
            $context.SaveVendorText($credentialsFileName, "")
        }
    }

    # Retrieve the current PowerShell edition
    $currentPowerShellEdition = $PSVersionTable.PSEdition

    Context "when the PowerShell edition is not valid" {
        # Declare the task inputs
        $taskInputs = @{
            applicationId               = "Application ID"
            clientSecret                = "Client Secret"
            tenantDomain                = "Tenant Domain"
            verifyAgainstMicrosoftGraph = "Yes"
        }

        # Modify the PowerShell edition
        $PSVersionTable.PSEdition = "Core"

        It "outputs a warning" {
            # Call the task
            Invoke-Expression ". $($taskScriptPath) -credentialsName 'SaveAzureADApplicationCredentialsTest' @taskInputs" -WarningVariable warningVariable

            # Verify the outputs
            $warningVariable | Should Not BeNullOrEmpty
        }
    }

    # Set the PowerShell edition back
    $PSVersionTable.PSEdition = $currentPowerShellEdition

    # Retrieve the current PowerShell version
    $currentPowerShellVersion = $PSVersionTable.PSVersion

    Context "when the PowerShell version is not valid" {
        # Declare the task inputs
        $taskInputs = @{
            applicationId               = "Application ID"
            clientSecret                = "Client Secret"
            tenantDomain                = "Tenant Domain"
            verifyAgainstMicrosoftGraph = "Yes"
        }

        # Modify the PowerShell version
        $PSVersionTable.PSVersion = [Version]"5.0"

        It "outputs a warning" {
            # Call the task
            Invoke-Expression ". $($taskScriptPath) -credentialsName 'SaveAzureADApplicationCredentialsTest' @taskInputs" -WarningVariable warningVariable

            # Verify the outputs
            $warningVariable | Should Not BeNullOrEmpty
        }
    }

    # Set the PowerShell version back
    $PSVersionTable.PSVersion = $currentPowerShellVersion

    Context "when the credentials are not selected to be verified against Microsoft Graph" {
        # Declare mocks
        Mock Get-MicrosoftGraphAuthenticationToken {}

        It "skips the verification" {
            # Declare the task inputs
            $taskInputs = @{
                applicationId               = "Application ID"
                clientSecret                = "Client Secret"
                tenantDomain                = "Tenant Domain"
                verifyAgainstMicrosoftGraph = "No"
            }

            # Call the task script
            Invoke-Expression ". $($taskScriptPath) -credentialsName 'SaveAzureADApplicationCredentialsTest' @taskInputs" -WarningVariable warningVariable

            # Verify the mocks
            Assert-MockCalled Get-MicrosoftGraphAuthenticationToken -Times 0 -Exactly -Scope It

            # Verify the output
            $warningVariable | Should BeNullOrEmpty
            $context.Outputs.errorMessages | Should BeNullOrEmpty
            $context.Outputs.verificationAgainstMicrosoftGraphResult | Should Be "Not run"
        }
    }

    Context "when a null or empty Microsoft Graph authentication token is returned" {
        # Declare mocks
        Mock Get-MicrosoftGraphAuthenticationToken {}

        It "outputs a warning" {
            # Declare the task inputs
            $taskInputs = @{
                applicationId               = "Application ID"
                clientSecret                = "Client Secret"
                tenantDomain                = "Tenant Domain"
                verifyAgainstMicrosoftGraph = "Yes"
            }

            # Call the task script
            Invoke-Expression ". $($taskScriptPath) -credentialsName 'SaveAzureADApplicationCredentialsTest' @taskInputs" -WarningVariable warningVariable

            # Verify the output
            $warningVariable | Should Not BeNullOrEmpty
            $context.Outputs.errorMessages | Should Not BeNullOrEmpty
            $context.Outputs.verificationAgainstMicrosoftGraphResult | Should Be "Failed"
        }
    }

    Context "when an exception occurs while retrieving the Microsoft Graph authentication token" {
        # Declare mocks
        Mock Get-MicrosoftGraphAuthenticationToken {
            throw "exception"
        }

        It "outputs a warning" {
            # Declare the task inputs
            $taskInputs = @{
                applicationId               = "Application ID"
                clientSecret                = "Client Secret"
                tenantDomain                = "Tenant Domain"
                verifyAgainstMicrosoftGraph = "Yes"
            }

            # Call the task script
            Invoke-Expression ". $($taskScriptPath) -credentialsName 'SaveAzureADApplicationCredentialsTest' @taskInputs" -WarningVariable warningVariable

            # Verify the output
            $warningVariable | Should Not BeNullOrEmpty
            $context.Outputs.errorMessages | Should Not BeNullOrEmpty
            $context.Outputs.verificationAgainstMicrosoftGraphResult | Should Be "Failed"
        }
    }

    Context "when a function unexpectedly throws an exception" {
        # Declare mocks
        Mock Get-MicrosoftGraphAuthenticationToken {
            return "token"
        }
        Mock Write-Information {
            throw "exception"
        }

        It "outputs a warning" {
            # Declare the task inputs
            $taskInputs = @{
                applicationId               = "Application ID"
                clientSecret                = "Client Secret"
                tenantDomain                = "Tenant Domain"
                verifyAgainstMicrosoftGraph = "Yes"
            }

            # Call the task script
            Invoke-Expression ". $($taskScriptPath) -credentialsName 'SaveAzureADApplicationCredentialsTest' @taskInputs" -WarningVariable warningVariable

            # Verify the output
            $warningVariable | Should Not BeNullOrEmpty
            $context.Outputs.errorMessages | Should Not BeNullOrEmpty
        }

        It "outputs an error if the flag to halt execution is set" {
            # Declare the task inputs
            $taskInputs = @{
                applicationId               = "Application ID"
                clientSecret                = "Client Secret"
                tenantDomain                = "Tenant Domain"
                verifyAgainstMicrosoftGraph = "Yes"
                haltExecutionOnError        = "Yes"
            }

            # Call the task script
            Invoke-Expression ". $($taskScriptPath) -credentialsName 'SaveAzureADApplicationCredentialsTest' @taskInputs" -ErrorVariable errorVariable

            # Verify the output
            $errorVariable | Should Not BeNullOrEmpty
            $context.Outputs.errorMessages | Should Not BeNullOrEmpty
        }
    }
}