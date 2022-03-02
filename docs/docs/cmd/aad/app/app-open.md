# aad app open

Gets the url of the Azure AD app registration page to be able to quickly navigate to the Azure Portal. 

## Usage

```sh
m365 aad app open [options]
```

## Options

`--appId <appId>`
: Application (client) ID of the Azure AD application registration to get.

`--preview`
: Use to get the url of the Azure AD preview portal.

`--autoOpenBrowser`
: Use to automatically open the url in the browser.

--8<-- "docs/cmd/_global.md"

## Examples

Prints the url of the Azure AD application registration on the Azure Portal

```sh
m365 aad app open --appId d75be2e1-0204-4f95-857d-51a37cf40be8
```

Prints the url of the Azure AD application registration on the preview Azure Portal

```sh
m365 aad app open --appId d75be2e1-0204-4f95-857d-51a37cf40be8 --preview
```

Opens the url in the browser

```sh
m365 aad app open --appId d75be2e1-0204-4f95-857d-51a37cf40be8 --autoOpenBrowser
```
