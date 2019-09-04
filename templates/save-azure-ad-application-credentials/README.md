## Description

This template saves a set of Azure AD application credentials to a Voleer workspace, allowing it to be used in other templates.

The set of credentials can also be verified against Microsoft Graph before saving it. If the verification fails, the credentials will not be saved.

## Inputs

Credential Name

   - This name will be used to load this set of credentials in the future.

Application ID

   - The application ID to save.

Client Secret

   - The client secret to save.

Tenant Domain

   - The domain of the tenant.

Verify Credentials?

   - Selects whether to verify the credentials against Microsoft Graph.

## Additional Notes

To verify that the credentials have been verified and saved successfully, or in the event that execution failed, logs can be downloaded by clicking on 'Download logs' in the 'Save Azure AD Application Credentials' step under the Activity Log.

If the credentials were saved successfully, the logs will contain an output message 'The provided credentials have been saved as CREDENTIAL_NAME' for future reference.
