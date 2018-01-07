# spo customaction list

Lists user custom actions for site or site collection

## Usage

```sh
spo customaction list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|Url of the site or site collection to retrieve the custom action from
`-s, --scope [scope]`|Scope of the custom action. Allowed values `Site|Web|All`. Default `All`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To retrieve list of custom actions, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Return details about all user custom actions located in site or site collection _https://contoso.sharepoint.com/sites/test_

```sh
spo customaction list -u https://contoso.sharepoint.com/sites/test
```

Return details about all user custom actions located in site or site collection _https://contoso.sharepoint.com/sites/test_

```sh
spo customaction list --url https://contoso.sharepoint.com/sites/test
```

Return details about all user custom actions located in site collection _https://contoso.sharepoint.com/sites/test_

```sh
spo customaction list -u https://contoso.sharepoint.com/sites/test -s Site
```

Return details about all user custom actions located in site _https://contoso.sharepoint.com/sites/test_

```sh
spo customaction list --url https://contoso.sharepoint.com/sites/test --scope Web
```

## More information

- UserCustomAction REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction](https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction)