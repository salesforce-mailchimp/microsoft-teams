function Global:Get-SavedCredentials {
    [CmdletBinding(PositionalBinding=$false)]
    [OutputType([PSObject])]
    param (
        [Parameter(Mandatory=$true, ParameterSetName="UserCredentials")]
        [Parameter(Mandatory=$true, ParameterSetName="ApplicationCredentials")]
        [ValidateNotNullOrEmpty()]
        [String]$CredentialsName,

        [Parameter(Mandatory=$true, ParameterSetName="UserCredentials")]
        [ValidateNotNull()]
        [Switch]$UserCredentials,

        [Parameter(Mandatory=$true, ParameterSetName="ApplicationCredentials")]
        [ValidateNotNull()]
        [Switch]$ApplicationCredentials
    )

    # Retrieve user credentials
    if ($UserCredentials) {
        # Check the vendor scope for the credentials
        $credentialsFileContents = $context.GetVendorText("AzureADUserCredentials_$($CredentialsName).json")
        if (![String]::IsNullOrWhiteSpace($credentialsFileContents)) {
            return $credentialsFileContents | ConvertFrom-Json
        }

        # Check the package scope for the credentials
        else {
            $credentialsFileContents = $context.GetPackageText("Office365_$($CredentialsName).json")
            if (![String]::IsNullOrWhiteSpace($credentialsFileContents)) {
                # Migrate the credentials from the package scope to the vendor scope
                $context.SaveVendorText("AzureADUserCredentials_$($CredentialsName).json", $credentialsFileContents)
                return $credentialsFileContents | ConvertFrom-Json
            }
        }
    }

    # Retrieve application credentials
    else {
        # Check the vendor scope for the credentials
        $credentialsFileContents = $context.GetVendorText("AzureADApplicationCredentials_$($CredentialsName).json")
        if (![String]::IsNullOrWhiteSpace($credentialsFileContents)) {
            return $credentialsFileContents | ConvertFrom-Json
        }

        # Check the package scope for the credentials
        else {
            $credentialsFileContents = $context.GetPackageText("MicrosoftGraph_$($CredentialsName).json")
            if (![String]::IsNullOrWhiteSpace($credentialsFileContents)) {
                # Migrate the credentials from the package scope to the vendor scope
                $context.SaveVendorText("AzureADApplicationCredentials_$($CredentialsName).json", $credentialsFileContents)
                return $credentialsFileContents | ConvertFrom-Json
            }
        }
    }
}