# spo customaction clear

Deletes all custom actions from site or site collection

## Usage

```sh
m365 spo customaction clear [options]
```

## Options

`-u, --url <url>`
: Url of the site or site collection to clear the custom actions from

`-s, --scope [scope]`
: Scope of the custom action. Allowed values `Site,Web,All`. Default `All`

`--confirm`
: Don't prompt for confirming removing all custom actions

--8<-- "docs/cmd/_global.md"

## Examples

Clears all user custom actions for both site and site collection _https://contoso.sharepoint.com/sites/test_.
Skips the confirmation prompt message.

```sh
m365 spo customaction clear --url https://contoso.sharepoint.com/sites/test --confirm
```

Clears all user custom actions for site _https://contoso.sharepoint.com/sites/test_. 

```sh
m365 spo customaction clear --url https://contoso.sharepoint.com/sites/test --scope Web
```

Clears all user custom actions for site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction clear --url https://contoso.sharepoint.com/sites/test --scope Site
```