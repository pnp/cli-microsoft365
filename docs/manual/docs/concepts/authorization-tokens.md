# Authorization and access tokens

Commands provided with the Office 365 CLI manipulate different settings of Office 365. Before you can execute any of the commands in the CLI, you have to connect to the Office 365 service corresponding to your command, such as SharePoint or Microsoft Graph.

## TL;DR

Use `Auth.ensureAccessToken` when:

- the URL of the service to which you are connected and the URL of the API you're communicating with are the same, eg. you're connected to AAD Graph at `https://graph.windows.net` and you're calling `https://graph.windows.net/myorganization/servicePrincipals` or you're connected to SharePoint tenant admin site at `https://contoso-admin.sharepoint.com` and you're calling the tenant admin API at `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`
- you want to renew the refresh token persisted in the Office 365 CLI

Use `Auth.getAccessToken` when:

- the URL of the service to which you are connected and the URL of the API you're calling are different, eg. you're connected to a SharePoint Online tenant admin site at `https://contoso-admin.sharepoint.com`, but you're calling `https://contoso.sharepoint.com/sites/team/_api/web/tenantappcatalog/AvailableApps/GetById('cd807889-9107-416e-88e4-7575789ab35c')/install` to install an app on a regular SharePoint site
- you're obtaining access token for a service different than the one to which you are connected and don't want to break the existing connection

## Authorization in the Office 365 CLI

There are a number of ways in which you can authenticate and authorize with Office 365. The Office 365 CLI uses the OAuth protocol to authorize with Office 365 and its services. OAuth flows in Office 365 are facilitated by Azure Active Directory.

### Azure AD application used by the Office 365 CLI

Office 365 CLI gets access to Office 365 through a custom Azure AD application named _PnP Office 365 Management Shell_. If you don't want to consent this application in your tenant, you can use a different application instead.

When specifying a custom Azure AD application to be used by the Office 365 CLI, you can either choose to use one application for all Office 365 services or a separate application for each service. To use one Azure AD application for all Office 365 CLI commands, set the `OFFICE365CLI_AADAPPID` environment variable to the ID of your Azure AD application. If you want to use a different Azure AD application for each Office 365 service use the following environment variables:

- `OFFICE365CLI_AADAADAPPID` - for the ID of the Azure AD application to communicate with Azure AD Graph
- `OFFICE365CLI_AADAZMGMTAPPID` - for the ID of the Azure AD application to communicate with the Azure Management Service
- `OFFICE365CLI_AADGRAPHAPPID` - for the ID of the Azure AD application to communicate with the Microsoft Graph
- `OFFICE365CLI_AADSPOAPPID` - for the ID of the Azure AD application to communicate with SharePoint Online

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
- Windows Azure Active Directory
    - Access the directory as the signed-in user
- Windows Azure Service Management API
    - Access Azure Service Management as organization users

!!! attention
    After changing the ID of the Azure AD application used by the Office 365 CLI refresh the existing connections to Office 365 using the corresponding `<service> connect` command. If you try to use the existing connection, Office 365 CLI will fail when trying to refresh the existing access token.

### Access and refresh tokens in the Office 365 CLI

After completing the OAuth flow, the CLI receives from Azure Active Directory a refresh- and an access token. Each web request to Office 365 APIs contains the access token which authorizes the Office 365 CLI to execute the particular operation. When the access token expires, the CLI uses the refresh token to obtain a new access token. When the refresh token expires, the user has to reconnect to Office 365 to obtain a new refresh token.

## Services and commands

Each command in the Office 365 CLI belongs to a service, for example the [spo site add](../cmd/spo/site/site-add.md) command, which creates a new modern site, belongs to the SharePoint Online service, while the [aad sp get](../cmd/aad/sp/sp-get.md) command, which lists Azure Active Directory service principals, belongs to the Azure Active Directory Graph service. Each service in Office 365 is a different Azure Active Directory authorization resource and requires a separate access token. When working with the CLI, you can be simultaneously connected to multiple services. Each command in the CLI knows which Office 365 service it communicates with and for which resource it should have a valid access token.

## Communicating with Office 365

Before a command can connect to Office 365, it requires a valid access token. Office 365 CLI offers you two methods to obtain a valid access token: [Auth.ensureAccessToken](https://github.com/SharePoint/office365-cli/blob/8b4ede874923fbe5fd84ebe79dc20206da18a529/src/Auth.ts#L62-L214) and [Auth.getAccessToken](https://github.com/SharePoint/office365-cli/blob/8b4ede874923fbe5fd84ebe79dc20206da18a529/src/Auth.ts#L216-L255). While the methods seem similar, they work differently and are meant for different purposes.

### Refresh access token for the current resource

If you're building a command that operates on the same URL, as the service to which the user is connected, you should use the `Auth.ensureAccessToken` method to refresh the token. Not only does this method resolve to a valid access token, which you can use directly in your code, but also stores the new access- and refresh token in the `auth.service` object from which you can use in any point in code.

For example: you're building a command that retrieves the list of service principals from the Azure Active Directory ([aad sp get](../cmd/aad/sp/sp-get.md)). The command uses the Azure Active Directory Graph API (`https://graph.windows.net`) for this. Since the resource URL of the AAD Graph service (`https://graph.windows.net`) and the URL of the API that the command has to call (`https://graph.windows.net/myorganization/servicePrincipals`) are both located on the same domain `https://graph.windows.net`, you should call the `Auth.ensureAccessToken` method to obtain a valid access token for the AAD Graph service, before calling the API.

As another example, let's take a method that communicates with the SharePoint Online tenant admin API to set a tenant property ([spo storageentity set](../cmd/spo/storageentity/storageentity-set.md)). While using the [spo connect](../cmd/spo/connect.md) command, users can connect to any SharePoint site, the _spo storageentity set_ command requires the user to be connected to the SharePoint tenant admin site. As a result, both the URL of the service, to which the user is connected (`https://contoso-admin.sharepoint.com`) and the URL of the API used by the service (`https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) are located on the same domain `https://contoso-admin.sharepoint.com`, which is why you should call the `Auth.ensureAccessToken` method to obtain a valid access for the SharePoint Online tenant admin site, before calling the API.

### Obtain a valid access token for a different resource

In some cases, while the user is connected to one service (with a corresponding resource), you need to retrieve a valid access token for a different resource. This is often the case when building SharePoint commands. In Office 365, SharePoint resources are spread over the following different domains: `contoso-admin.sharepoint.com` - which hosts the SharePoint Online tenant admin site, `contoso.sharepoint.com` - which hosts regular SharePoint sites and `contoso-my.sharepoint.com` - which hosts users' OneDrive sites. From Azure Active Directory point of view, all these domains are perceived as separate resources and require different access tokens.

If you're building a command that allows users to specify a URL on which the operation should be performed, such as [spo app install](../cmd/spo/app/app-install.md), you cannot make any assumptions of the service to which they're connected, and you should obtain a valid access token for the resources corresponding to the URL specified by the user, using the `Auth.getResourceToken` method. While the `Auth.getResourceToken` method also returns a valid access token for the specified resource, it doesn't update the connection information on the `auth.service` object.

## Why two different methods to get tokens

In the Office 365 CLI, connecting to Office 365 services is interactive and requires user input. If commands relied on context information from the site, to which users are connected, it would be impossible to build scripts using the Office 365 CLI commands.

Some SharePoint commands require to be executed in the context of the tenant admin site. Some commands additionally require specifying tenant information. For performance reasons, this information is retrieved only initially, when connecting to SharePoint, if the specified site to connect to is a tenant admin site. Because tenant information doesn't change, there is no point in retrieving it on every call to SharePoint. If users would first connect to a regular SharePoint site, and would then switch to the tenant admin site, the tenant information would be missing and the commands requiring it would fail.

## Rules of thumb

- if you're building a command for a service other than SharePoint Online, you will most likely use the `Auth.ensureAccessToken` method
- if the command you're building allows the user to specify a URL on which the API is called, you should use the `Auth.getAccessToken` method to get access token for the URL specified by the user