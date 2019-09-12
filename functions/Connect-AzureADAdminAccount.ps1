<#
.SYNOPSIS
    This function connects to Azure AD using admin account credentials.
.DESCRIPTION
    This function connects to Azure AD using admin account credentials.
    It returns whether the connection and logon was successful.
#>
function Connect-AzureADAdminAccount {
    param (
        # The username of the AzureAD account.
        [Parameter(Mandatory=$true)]
        [String]$Username,

        # The password of the AzureAD account.
        [Parameter(Mandatory=$true)]
        [SecureString]$Password
    )

    # Create the AzureAD credential from the given username and password
    $azureADCredential = [PSCredential]::new($Username, $Password)

    # Logon to AzureAD
    try {
        Connect-AzureAD -Credential $azureADCredential -ErrorAction Stop

        # Logon was successful
        Write-Information "Connection and logon to Azure AD successful with username '$($username)'."
        return $true
    }

    # Logon was unsuccessful
    catch {
        Write-Error "Failed Azure AD account login with username '$($username)'.`r`n$($_.Exception.Message)"
        return $false
    }
}
