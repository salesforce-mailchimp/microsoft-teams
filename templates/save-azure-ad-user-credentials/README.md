## Description

This template saves a set of Azure AD user credentials to a Voleer workspace, allowing it to be used in other templates.

The set of credentials can also be verified for logins against various Microsoft services, before saving it. If any of the services selected for verification fail, the credentials will not be saved.

Visit the [Voleer Videos](https://www.voleer.com/videos) page for a video walkthrough on how to use this template.

## Inputs

Credential Name

   - The credentials will be saved under this name, and the same name will be used to load this set of credentials in the future.

Username

   - The username to save.

Password

   - The password to save.

Verify Login against Azure?

   - Selects whether to verify the credentials provided for login against Azure.

Verify Login against Azure AD?

   - Selects whether to verify the credentials provided for login against Azure AD.

Verify Login against Azure RM?

   - Selects whether to verify the credentials provided for login against Azure RM.

Verify Login against Exchange Online?

   - Selects whether to verify the credentials provided for login against Exchange Online.

Verify Login against MS Online?

   - Selects whether to verify the credentials provided for login against MS Online.

Verify Login against Office 365 Security and Compliance Center?

   - Selects whether to verify the credentials provided for login against Office 365 Security and Compliance Center.

## Additional Notes

The verification steps will not be successful if the user account has multi-factor authentication (MFA) enabled.

To verify that the credentials have been verified and saved successfully, or in the event that execution failed, logs can be downloaded by clicking on 'Download logs' in the 'Save Azure AD User Credentials' step under the Activity Log for this template.

If the credentials were saved successfully, the logs will contain an output message 'The provided credentials have been saved as CREDENTIAL_NAME' for future reference.
