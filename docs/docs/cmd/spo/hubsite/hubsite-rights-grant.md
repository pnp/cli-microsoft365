# spo hubsite rights grant

Grants permissions to join the hub site for one or more principals

## Usage

```sh
m365 spo hubsite rights grant [options]
```

## Options

`-u, --hubSiteUrl <hubSiteUrl>`
: The URL of the hub site to grant rights on

`-p, --principals <principals>`
: Comma-separated list of principals to grant join rights. Principals can be users or mail-enabled security groups in the form of `alias` or `alias@<domain name>.com`

`-r, --rights <rights>`
: Rights to grant to principals. Available values `Join`

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global or SharePoint administrator.

## Examples

Grant user with alias _PattiF_ permission to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_

```sh
m365 spo hubsite rights grant --hubSiteUrl https://contoso.sharepoint.com/sites/sales --principals PattiF --rights Join
```

Grant users with aliases _PattiF_ and _AdeleV_ permission to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_

```sh
m365 spo hubsite rights grant --hubSiteUrl https://contoso.sharepoint.com/sites/sales --principals "PattiF,AdeleV" --rights Join
```

Grant user with email _PattiF@contoso.com_ permission to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_

```sh
m365 spo hubsite rights grant --hubSiteUrl https://contoso.sharepoint.com/sites/sales --principals PattiF@contoso.com --rights Join
```

## Response

The command won't return a response on success.

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
