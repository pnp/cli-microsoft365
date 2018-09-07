# Logging in to Office 365

Before you can use Office 365 CLI commands to manage your tenant, you have to log in to Office 365. Following section explains how you can log in and check the Office 365 login status.

## Office 365 services

Using the Office 365 CLI you can manage different areas of an Office 365 tenant. Currently, commands for working with SharePoint Online, Azure Active Directory, Microsoft Graph and the Azure Management Service are available, but more commands for other services will be added in the future.

Commands in the Office 365 CLI are organized into services. For example, all commands that manage SharePoint Online begin with `spo` (`spo app list`, `spo cdn get`, etc.) and commands for working with the Azure AD begin with `aad`. For each Office 365 service, the CLI offers three commands for managing the login with that service.

### Log in to an Office 365 service

Office 365 CLI offers you a number of ways to log in to Office 365 services in your tenant.

#### Log in using the default device code flow

The default way to log in to an Office 365 service using the Office 365 CLI is through the device code flow. To log in to an Office 365 service, use the `<service> login` command for that service. For example, to log in to SharePoint Online, execute:

```sh
spo login https://contoso.sharepoint.com
```

To log in to Azure AD, which uses a fixed URL, execute:

```sh
aad login
```

!!! tip
    If the service uses a fixed URL, such as Azure AD or Microsoft Graph, you will execute the `login` command without any arguments, for example `aad login`. However, when logging in to other services that require a URL, such as SharePoint, you will execute the `login` command with the URL to which the CLI should log in to, for example: `spo login https://contoso.sharepoint.com`. For more information on logging in to each service, refer to the help of the `login` command for that service.

After executing the `login` command, you will be prompted to navigate to _https://aka.ms/devicelogin_ in your web browser and enter the login code presented to you by the Office 365 CLI in the command line. After entering the code, you will see the prompt that you are about to authenticate the _PnP Office 365 Management Shell_ application to access your tenant on your behalf.

[![Signing in to Azure Active Directory](../images/login.png)](../images/login.png)

If you are using the Office 365 CLI for the first time, you will be also prompted to verify the permissions you are about to grant the Office 365 CLI. This is referred to as _consent_.

[![Granting the Office 365 CLI the necessary permissions](../images/consent.png)](../images/consent.png)

The device code flow is the recommended approach for command-line tools to authenticate with resources secured with Azure AD. Because the authentication process is handled in the browser by Azure AD itself, it allows you to benefit of rich security features such as multi-factor authentication or conditional access. The device code flow is interactive and requires user interaction which might be limiting if you want to use the Office 365 CLI in your continuous deployment setup which is fully automated and doesn't involve user interaction.

#### Log in using user name and password

An alternative way to log in to an Office 365 service in the Office 365 CLI is by using a user name and password. To use this way of authenticating, set the `authType` option to `password` and specify your credentials using the `userName` and `password` options.

To log in to SharePoint Online using your user name and password, execute:

```sh
spo login https://contoso.sharepoint.com --authType password --userName user@contoso.com --password pass@word1
```

To log in to Azure AD using your user name and password, execute:

```sh
aad login --authType password --userName user@contoso.com --password pass@word1
```

Using credentials to log in to Office 365 is convenient in automation scenarios where you cannot authenticate interactively. The downside of this way of authenticating is, that it doesn't allow you to use any of the advanced security features that Azure AD offers. If your account for example uses multi-factor authentication, logging in to Office 365 using credentials will fail.

!!! attention
    When logging in to Office 365 using credentials, Office 365 CLI will persist not only the retrieved access and refresh token, but also the credentials you specified when logging in. This is necessary for the CLI to be able to retrieve a new refresh token, in case the previously retrieved refresh token expired or has been invalidated.

Generally, you should use the default device code flow. If you need to use a non-interactive authentication flow, you can authenticate using credentials of an account that has sufficient privileges in your tenant and doesn't have multi-factor authentication or other advanced security features enabled.

### Check login status

To see if you're logged in to the particular Office 365 service and if so, with which account, use the `<service> status` command, for example, to see if you're logged in to SharePoint Online, execute:

```sh
spo status
```

### Log out from an Office 365 service

To log out from an Office 365 service, use the `<service> logout` command for that service. For example, to log out from SharePoint Online, execute:

```sh
spo logout
```

!!! tip
    Each service in the Office 365 CLI manages its login information independently. This makes it possible for you to be logged in to different services with different accounts. Using the `<service> status` command you can see which account is currently logged in to the particular service.

<script src="https://asciinema.org/a/158294.js" id="asciicast-158294" async></script>

### Logging in to SharePoint Online

When logging in to SharePoint Online, you can log in either to the tenant admin site (eg. `https://contoso-admin.sharepoint.com`) or any other site in your tenant. If you are logged in to the tenant admin site, but would like to get information for some other site, such as the list of its subsites or lists, the CLI will automatically switch to that site, without you having to reauthenticate.

!!! attention
    Please note, that some commands require login to the tenant admin site, and if you try to execute them, while being logged in to a different site, you will get an error. For more information whether the login to the tenant admin site is required or not, refer to the help of that particular command.

!!! tip
    The most convenient way of working with the CLI is to log in to the tenant admin site. Based on the options you specified when executing commands, the CLI will automatically switch between the tenant admin site or other sites that you will want to manage. You should log in to other sites only, if you don't have tenant admin privileges and yet would like to automate some of your work using the Office 365 CLI.

## Authorize with Office 365

To authorize for communicating with Office 365 API, the Office 365 CLI uses the OAuth 2.0 protocol. When using the default device code flow, you authenticate with Azure AD in the web browser. After authenticating, Office 365 CLI will attempt to retrieve a valid access token for the specified Office 365 service. If you have insufficient permissions to access the particular service, authorization will fail with an adequate error.

If you authenticate using credentials, the authentication and authorization are a part of the same request that Office 365 CLI issues towards Azure AD. If either authentication or authorization fails, you will see a corresponding error message explaining what went wrong.

## Re-consent the PnP Office 365 Management Shell Azure AD application

Office 365 CLI uses the _PnP Office 365 Management Shell_ Azure AD application to log in to your Office 365 tenant on your behalf. As we add new commands to the CLI, it's possible, that new permissions will be added to the _PnP Office 365 Management Shell_ Azure AD application. To be able to use the newly added commands which depend on these new permissions, you will have to re-approve the _PnP Office 365 Management Shell_ Azure AD application in your Azure AD. This process is also known as _re-consenting the Azure AD application_.

To re-consent the _PnP Office 365 Management Shell_ application in your Azure AD, in the command line execute:

```sh
o365 --reconsent
```

Office 365 CLI will provide you with a URL that you should open in the web browser and sign in with your organizational account. After authenticating, Azure AD will prompt you to approve the new set of permissions. Once you approved the permissions, you will be redirected to an empty page. You can now use all commands in the Office 365 CLI.

## Logging in to Office 365 via a proxy

All communication between the Office 365 CLI and Office 365 APIs happens via web requests. If you're behind a proxy, you should set up an environment variable to allow Office 365 CLI to log in to Office 365. More information about the necessary configuration steps is available at [https://github.com/request/request#controlling-proxy-behaviour-using-environment-variables](https://github.com/request/request#controlling-proxy-behaviour-using-environment-variables).

## Persisted connections

After logging in to Office 365, the Office 365 CLI will persist that connection information until you explicitly log out from the particular service. This is necessary to support building scripts using the Office 365 CLI, where each command is executed independently of other commands. Persisted connection contains information about the user name used to establish the connection, the connected Office 365 service URL, the access token and the refresh token. To secure this information from unprivileged access, it's stored securely in the password store specific to the platform on which you're using the CLI. For more information, see the separate article dedicated to [persisting connection information](../concepts/persisting-connection.md) in the Office 365 CLI.