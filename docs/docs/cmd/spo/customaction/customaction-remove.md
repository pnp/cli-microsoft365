# spo customaction remove

Removes specified custom action from site or site collection

## Usage

```sh
m365 spo customaction remove [options]
```

## Options

`-i, --id [id]`
: Id (GUID) of the custom action to remove. Specify either `id` or `title`

`-t, --title [title]`
: Title of the user custom action to retrieve information for. Specify either `id` or `title`

`-u, --url <url>`
: Url of the site or site collection to remove the custom action from

`-s, --scope [scope]`
: Scope of the custom action. Allowed values `Site,Web,All`. Default `All`

`--confirm`
: Don't prompt for confirming removal of a user custom action

--8<-- "docs/cmd/_global.md"

## Remarks

If the command finds multiple user custom actions with the specified title, it will prompt you to disambiguate which user custom action it should use, listing the discovered IDs.

## Examples

Removes user custom action with ID _058140e3-0e37-44fc-a1d3-79c487d371a3_ located in site or site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction remove --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test
```

Removes user custom action with Title _YourAppCustomizer_ located in site or site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction remove --title "YourAppCustomizer" --url https://contoso.sharepoint.com/sites/test
```

Removes user custom action with ID _058140e3-0e37-44fc-a1d3-79c487d371a3_ located in site or site collection _https://contoso.sharepoint.com/sites/test_. Skips the confirmation prompt message.

```sh
m365 spo customaction remove --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test --confirm
```

Removes user custom action with Title _YourAppCustomizer_ located in site or site collection _https://contoso.sharepoint.com/sites/test_. Skips the confirmation prompt message.

```sh
m365 spo customaction remove --title "YourAppCustomizer" --url https://contoso.sharepoint.com/sites/test --confirm
```

Removes user custom action with ID _058140e3-0e37-44fc-a1d3-79c487d371a3_ located in site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction remove --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test --scope Site
```

Removes user custom action with Title _YourAppCustomizer_ located in site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction remove --title "YourAppCustomizer" --url https://contoso.sharepoint.com/sites/test --scope Site
```

Removes user custom action with ID _058140e3-0e37-44fc-a1d3-79c487d371a3_ located in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction remove --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test --scope Web
```

Removes user custom action with Title _YourAppCustomizer_ located in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction remove --title "YourAppCustomizer" --url https://contoso.sharepoint.com/sites/test --scope Web
```

## More information

- UserCustomAction REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction](https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction)
