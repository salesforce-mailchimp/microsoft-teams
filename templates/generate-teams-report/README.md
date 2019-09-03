## Description

- This template imports and tags contacts from Salesforce into Mailchimp using a Salesforce query.

- By using this template, you will be able to use the rich Salesforce query builder to create more targeted campaigns within Mailchimp.

## Inputs

Salesforce Credentials Name

   - This name is used to load a set of credentials, which will be used to access Salesforce.
   - The credentials can be saved by using the `Save Salesforce Credentials` template.

Mailchimp Credentials Name

   - This name is used to load a set of credentials, which will be used to access Mailchimp.
   - The credentials can be saved by using the `Save Mailchimp Credentials` template.

Query

   - The query to select the contacts to export from Salesforce
   - The query can be tested at this page <https://na112.salesforce.com/_ui/common/apex/debug/ApexCSIPage>
   - Example query : SELECT Name FROM Contact WHERE Contact.Instance_Executed__C >= 10

Salesforce To Mailchimp Mappings Csv

   - The variable mappings between Salesforce and Mailchimp in a csv string
   - Example csv (default) :

"SalesForceName","MailchimpName","Type"
"MailingCountry","Country","text"
"Voleer\_Registration\_Date\_\_c","Voleer Registration Date","date"
"Region\_\_c","Region","text"
"FirstName","First Name","text"
"LastName","Last Name","text"

   - The default mappings will be used if this field is left empty or "empty" is provided.

List Name

   - The name of the list in Mailchimp to import the contacts

Tag Name

   - The name of the tag to add to the imported Mailchimp contacts

## Additional Notes

If the tag is used as a trigger in an automated campaign, make sure the campaign is set up BEFORE
running this template. If a tag is added before it's associated with an automated campaign, the contacts
will NOT be triggered when the campaign is created.


