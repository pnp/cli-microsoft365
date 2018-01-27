# spo hubsite rights grant

Grants permissions to join the hub site for one or more principals

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

## Usage

```sh
spo hubsite rights grant [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|The URL of the hub site to grant rights on
`-p, --principals <principals>`|Comma-separated list of principals to grant join rights. Principals can be users or mail-enabled security groups in the form of `alias` or `alias@<domain name>.com`
`-r, --rights <rights>`|Rights to grant to principals. Available values `Join`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To grant permissions to join the hub site, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`. If you are connected to a different site and will try to grant permissions to join the hub site, you will get an error.

## Examples

Grant user with alias _PattiF_ permission to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_

```sh
spo hubsite rights grant --url https://contoso.sharepoint.com/sites/sales --principals PattiF --rights Join
```

Grant users with aliases _PattiF_ and _AdeleV_ permission to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_

```sh
spo hubsite rights grant --url https://contoso.sharepoint.com/sites/sales --principals PattiF,AdeleV --rights Join
```

Grant user with email _PattiF@contoso.com_ permission to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_

```sh
spo hubsite rights grant --url https://contoso.sharepoint.com/sites/sales --principals PattiF@contoso.com --rights Join
```

## More information

- SharePoint hub sites new in Office 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)