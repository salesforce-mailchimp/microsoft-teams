function Get-SavedCredentials {
    [CmdletBinding(PositionalBinding=$false)]
    [OutputType([PSObject])]
    param (
        [Parameter(Mandatory=$true, ParameterSetName="MailchimpCredentials")]
        [Parameter(Mandatory=$true, ParameterSetName="SalesforceCredentials")]
        [ValidateNotNullOrEmpty()]
        [String]$CredentialsName,

        [Parameter(Mandatory=$true, ParameterSetName="MailchimpCredentials")]
        [ValidateNotNull()]
        [Switch]$MailchimpCredentials,

        [Parameter(Mandatory=$true, ParameterSetName="SalesforceCredentials")]
        [ValidateNotNull()]
        [Switch]$SalesforceCredentials
    )

    # Retrieve Mailchimp credentials
    if ($MailchimpCredentials) {
        # Check the vendor scope for the credentials
        $credentialsFileContents = $context.GetVendorText("MailchimpCredentials_$($CredentialsName).json")
        if (![String]::IsNullOrWhiteSpace($credentialsFileContents)) {
            return $credentialsFileContents | ConvertFrom-Json
        }
    }

    # Retrieve Salesforce credentials
    else {
        # Check the vendor scope for the credentials
        $credentialsFileContents = $context.GetVendorText("SalesforceCredentials_$($CredentialsName).json")
        if (![String]::IsNullOrWhiteSpace($credentialsFileContents)) {
            return $credentialsFileContents | ConvertFrom-Json
        }
    }
}