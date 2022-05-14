# aad app open

Returns deep link of the current AD app to open the Azure portal on the Azure AD app registration management page.

## Usage

```sh
m365 app open [options]
```

## Options

`--appId [appId]`
: Optional Application (client) ID of the Azure AD application registration to open. Uses the app from the `.m365rc.json` file corresponding to the `appId`. If multiple apps are available, this will evade the prompt to choose an app. If the `appId` is not available in the list of apps, an error is thrown.

`--preview`
: Use to open the url of the Azure AD preview portal.

--8<-- "docs/cmd/_global.md"

## Remarks

If config setting `autoOpenLinksInBrowser` is configured to true, the command will automatically open the link to the Azure Portal in the browser.

Gets the app from the `.m365rc.json` file in the current directory. If the `--appId` option is not used and multiple apps are available, it will prompt the user to choose one.

## Examples

Prints the URL to the Azure AD application registration management page on the Azure Portal. 

```sh
m365 app open
```

Prints the url of the Azure AD application registration management page on the preview Azure Portal.

```sh
m365 app open --preview
```

Prints the URL to the Azure AD application registration management page on the Azure Portal, evading a possible choice prompt in the case of multiple saved apps in the `.m365rc.json` file. 

```sh
m365 app open --appId d75be2e1-0204-4f95-857d-51a37cf40be8 
```
