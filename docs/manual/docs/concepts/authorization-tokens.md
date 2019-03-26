# Authorization and access tokens

Commands provided with the Office 365 CLI manipulate different settings of Office 365. Before you can execute any of the commands in the CLI, you have to log in to Office 365. Office 365 CLI will then automatically retrieve the access token necessary to execute the particular command.

## Authorization in the Office 365 CLI

There are a number of ways in which you can authenticate and authorize with Office 365. The Office 365 CLI uses the OAuth protocol to authorize with Office 365 and its services. OAuth flows in Office 365 are facilitated by Azure Active Directory.

### Azure AD application used by the Office 365 CLI

Office 365 CLI gets access to Office 365 through a custom Azure AD application named _PnP Office 365 Management Shell_. If you don't want to consent this application in your tenant, you can use a different application instead.

!!! important
    When you decide to use your own Azure AD application, you need to choose the application to be a **public client**. Despite the setting's description, the application will not be publicly accessible. This setting enables the use of the device flow for your own application. Without activating this setting, it is not possible to complete the authentication process. The option is currently only available in the preview blade for managing for Azure AD applications.
    [![The 'public client' enabled for an Azure AD application](../images/activate-public-client-aad-app.png)](../images/activate-public-client-aad-app.png)

When specifying a custom Azure AD application to be used by the Office 365 CLI, set the `OFFICE365CLI_AADAPPID` environment variable to the ID of your Azure AD application.

Office 365 CLI requires the following permissions to Office 365 services:

- Office 365 SharePoint Online (Microsoft.SharePoint)
    - Have full control of all site collections
    - Read user profiles
    - Read and write managed metadata
- Microsoft Graph
    - Invite guest users to the organization
    - Read and write all groups
    - Read and write directory data
    - Access directory as the signed in user
    - Read and write identity providers
    - Send mail as a user
    - Read and write to all app catalogs
- Windows Azure Active Directory
    - Access the directory as the signed-in user
- Windows Azure Service Management API
    - Access Azure Service Management as organization users

!!! attention
    After changing the ID of the Azure AD application used by the Office 365 CLI refresh the existing connection to Office 365 using the `login` command. If you try to use the existing connection, Office 365 CLI will fail when trying to refresh the existing access token.

### Access and refresh tokens in the Office 365 CLI

After completing the OAuth flow, the CLI receives from Azure Active Directory a refresh- and an access token. Each web request to Office 365 APIs contains the access token which authorizes the Office 365 CLI to execute the particular operation. When the access token expires, the CLI uses the refresh token to obtain a new access token. When the refresh token expires, the user has to reauthenticate to Office 365 to obtain a new refresh token.

## Services and commands

Each command in the Office 365 CLI belongs to a service, for example the [spo site add](../cmd/spo/site/site-add.md) command, which creates a new modern site, belongs to the SharePoint Online service, while the [aad sp get](../cmd/aad/sp/sp-get.md) command, which lists Azure Active Directory service principals, belongs to the Azure Active Directory Graph service. Each service in Office 365 is a different Azure Active Directory authorization resource and requires a separate access token. When working with the CLI, you can be simultaneously connected to multiple services. Each command in the CLI knows which Office 365 service it communicates with and for which resource it should have a valid access token.

## Communicating with Office 365

Before a command can log in to Office 365, it requires a valid access token. Office 365 CLI automatically obtains the access token for the particular web request without you having to worry about it.