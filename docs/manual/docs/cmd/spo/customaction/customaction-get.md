# spo customaction get

Gets information about the specific user custom action for site or site collection

## Usage

```sh
spo customaction get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|ID of the app to retrieve information for
`-u, --url <url>`|Url of the site or site collection to retrieve the custom action from
`-s, --scope [scope]`|Scope of the custom action. Allowed values Site|Web|All. Default All
`--verbose`|Runs command with verbose logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To retrieve custom action, you have to first connect to a SharePoint Online site using the
[spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

```sh
spo customaction get -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/test
```

```sh
spo customaction get --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test
```

```sh
spo customaction get -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/test -s Site
```

```sh
spo customaction get --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test --scope Web
```

Returns details about the user custom action with ID 'b2307a39-e878-458b-bc90-03bc578531d6' available in the site or site collection.

## More information

- UserCustomAction REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction](https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction)