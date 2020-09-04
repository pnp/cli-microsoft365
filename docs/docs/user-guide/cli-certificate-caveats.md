# Caveats when working with the CLI and certificate login

## I get error "AADSTS700027 Client assertion contains an invalid signature" when I login the CLI with certificate, what am I doing wrong

There is an article ["Using your own Azure AD identity"](./using-own-identity.md) dedicated to using the CLI with your own identity, and you should have a look at it and see if it can help you. Many of the cases we've seen in the Github issues list are that people forget to set the `CLIMICROSOFT365_AADAPPID` or `CLIMICROSOFT365_TENANT` environment variables. Setting these variables could be as easy as adding them before your command on the bash command line like `CLIMICROSOFT365_AADAPPID=value1 CLIMICROSOFT365_TENANT=value2 m365 command` (see [#1532](https://github.com/pnp/cli-microsoft365/issues/1532) or [#1496](https://github.com/pnp/cli-microsoft365/issues/1496#issuecomment-625549739)). If you are Windows user the syntax should be like `set CLIMICROSOFT365_AADAPPID=value1` and `set CLIMICROSOFT365_TENANT=value2` then your cli command ([#1121](https://github.com/pnp/cli-microsoft365/issues/1121#issuecomment-533609882)).

## I get "Error: AADSTS700025: Client is public so 'client_assertion' should not be presented"

If you want to authenticate the CLI using and certificate, you shouldn't treat the application as a public client. You should set the default client type your Azure AD application to "NO." More information can be found in this issue [#948](https://github.com/pnp/cli-microsoft365/issues/948#issuecomment-487145809).

## What is the minimum set of Azure AD app permissions to execute SharePoint commands with a certificate CLI login

When you decide to use the CLI with your own Azure AD app to execute SharePoint CLI commands, you need to grant it at least the  Microsoft Graph `Sites.Read.All` permission, and then any other scopes required by the commands you'd like to execute. For example, if you'd like to list all the sites within your tenant using the `m365 spo site list` command, then the minimum permissions for your app would be Microsoft Graph `Sites.Read.All` and SharePoint `Sites.Manage.All`.

- your application requires the Microsoft Graph `Sites.Read.All` permission because the `m365 login` command is using `https://graph.microsoft.com/v1.0/sites/root?$select=webUrl` API to dynamically find the root SharePoint site URL and use it to get a token for the SharePoint resource from Microsoft Identity dynamically.
- your application requires the SharePoint `Sites.Manage.All` permission to list the sites since the CLI uses SharePoint APIs to do that from the `spo site list` command, and the minimum permissions to list sites is `Sites.Manage.All`.

[![Azure AD application permissions highlighted in Azure AD](../images/cli-certificate-caveats/min-app-permissions-to-list-SP-sites.png)](../images/cli-certificate-caveats/min-app-permissions-to-list-SP-sites.png)

Here is the result:

[![Result of running the m365 spo site list command](../images/cli-certificate-caveats/spo-list-sites-result.png)](../images/cli-certificate-caveats/spo-list-sites-result.png)

## I get an error: 403, "AccessDenied Either scp or roles claim need to be present in the token" when executing a CLI for Microsoft 365 SharePoint command. What does it mean

It means that the Azure AD application that the CLI is running under does not have Microsoft Graph `Sites.Read.All` application permission granted. If you are trying to use the CLI with a certificate login and SharePoint, you would have to allow Microsoft Graph `Sites.Read.All` application permissions to the Azure AD app.

## I am using CLI with a certificate, but when I execute the `spo site add` I get error "Insufficient privileges to complete the operation."

This error can occur when you use the CLI with a certificate login and try to create a new SharePoint Team site that uses Microsoft 365 group (WebTemplate: #GROUP). Getting this error is a known issue with the CLI and the SharePoint APIs, but there is a workaround. The workaround is to use the `m365 aad o365group add` command to create Team Sites.

If your goal is to create team sites, you can use the `m365 aad o365group add` command. The command is calling a Microsoft API that creates a Microsoft 365 group with a SharePoint site collection associated with the group.

Here is how to do it:
[![Arrow pointing from a modern site URL to the Microsoft 365 group's mail nickname](../images/cli-certificate-caveats/create-team-site-using-spo-o365group-add.png)](../images/cli-certificate-caveats/create-team-site-using-spo-o365group-add.png)

As I mentioned above, when creating a Microsoft 365 Group, the Microsoft 365 APIs create a site collection with it. The `mailNickname` property is the site URL of the site collection. You can combine the `m365 aad o365group add` command with `m365 spo site set` to change additional properties of the site not available in the `m365 aad o365group add` command. From the screenshot above, you can see that the `spo site set` command is used to change the site classification after having created the group and a SharePoint site. Combining these two commands will give you the same functionality as the `spo site add` command.

### What are the minimum permissions required to use the `m365 aad o365group add` command

You would need the Microsoft Graph `Group.Create` and `User.Read.All` application permissions.

[![Azure AD application permissions](../images/cli-certificate-caveats/min-app-permissions-create-m365group.png)](../images/cli-certificate-caveats/min-app-permissions-create-m365group.png)

### What are the minimum permissions required to use the `m365 aad o365group add` command and the `m365 spo site set` command

[![Azure AD application permissions](../images/cli-certificate-caveats/min-permissions-team-site.png)](../images/cli-certificate-caveats/min-permissions-team-site.png)

You would need the Microsoft Graph `Group.Create` and `User.Read.All` application permissions together with SharePoint `Sites.FullControl.All` application permission.

### Will the `spo site add` command and CLI certificate login work for creating Communication sites and Classic sites

Yes, it will. There is a known issue with creating modern Team sites, but the workaround above should sort that as well.

### Why not make the Team Sites being created by just executing `spo site add`

We are discussing this with the rest of the CLI team, and we might implement a fallback to the Microsoft Graph Group APIs to create the site in case of a CLI certificate login. So, it will use the same APIs as the `aad o365group add` command uses. Until this is implemented in the CLI we recommend to use the workaround described earlier.

### There is a well-documented API from Microsoft. Why does the CLI not use it to create the modern Team Sites

There is a well-documented API for the creation of modern sites indeed. Unfortunately, the document mentions that we cannot create a new Team site based on Microsoft 365 Group.

[![API limitations highlighted in the API docs](../images/cli-certificate-caveats/doc-not-apply-to-team-sites.png)](../images/cli-certificate-caveats/doc-not-apply-to-team-sites.png)
