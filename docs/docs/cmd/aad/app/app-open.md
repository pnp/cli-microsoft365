# aad app open

Returns deep link to open the Azure portal on the Azure AD app registration management page.

## Usage

```sh
m365 aad app open [options]
```

## Options

`--appId <appId>`
: Application (client) ID of the Azure AD application registration to open.

`--preview`
: Use to open the url of the Azure AD preview portal.

`--autoOpenBrowser`
: Use to automatically open the url in the browser.

--8<-- "docs/cmd/_global.md"

## Examples

Prints the url of the Azure AD application registration management page on the Azure Portal.

```sh
m365 aad app open --appId d75be2e1-0204-4f95-857d-51a37cf40be8
```

Prints the url of the Azure AD application registration management page on the preview Azure Portal.

```sh
m365 aad app open --appId d75be2e1-0204-4f95-857d-51a37cf40be8 --preview
```

Opens the url of the Azure AD application registration management page using the browser.

```sh
m365 aad app open --appId d75be2e1-0204-4f95-857d-51a37cf40be8 --autoOpenBrowser
```
