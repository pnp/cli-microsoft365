# app get

Retrieves information about the current Azure AD app

## Usage

```sh
m365 app get [options]
```

## Options

`--appId [appId]`
: Client ID of the Azure AD app registered in the .m365rc.json file to retrieve information for

--8<-- "docs/cmd/_global.md"

## Remarks

Use this command to quickly look up information for the Azure AD application registration registered in the .m365rc.json file in your current project (folder).

If you have multiple apps registered in your .m365rc.json file, you can specify the app for which you'd like to retrieve permissions using the `--appId` option. If you don't specify the app using the `--appId` option, you'll be prompted to select one of the applications from your .m365rc.json file.

## Examples

Retrieve information about your current Azure AD app

```sh
m365 app get
```

Retrieve information about the Azure AD app with client ID _e23d235c-fcdf-45d1-ac5f-24ab2ee0695d_ specified in the _.m365rc.json_ file

```sh
m365 app get --appId e23d235c-fcdf-45d1-ac5f-24ab2ee0695d
```
