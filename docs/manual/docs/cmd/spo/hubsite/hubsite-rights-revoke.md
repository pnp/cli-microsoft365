# spo hubsite rights revoke

Revokes rights to join sites to the specified hub site for one or more principals

## Usage

```sh
spo hubsite rights revoke [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|The URL of the hub site to revoke rights on
`-p, --principals <principals>`|Comma-separated list of principals to revoke join rights. Principals can be users or mail-enabled security groups in the form of `alias` or `alias@<domain name>.com`
`--confirm`|Don't prompt for confirming revoking rights
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To revoke rights to join sites to a hub site, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`. If you are connected to a different site and will try to revoke rights, you will get an error.

## Examples

Revoke rights to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_ from user with alias _PattiF_. Will prompt for confirmation before revoking the rights

```sh
spo hubsite rights revoke --url https://contoso.sharepoint.com/sites/sales --principals PattiF
```

Revoke rights to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_ from user with aliases _PattiF_ and _AdeleV_ without prompting for confirmation

```sh
spo hubsite rights revoke --url https://contoso.sharepoint.com/sites/sales --principals PattiF,AdeleV --confirm
```

## More information

- SharePoint hub sites new in Office 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)