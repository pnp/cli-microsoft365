# spo hubsite rights revoke

Revokes rights to join sites to the specified hub site for one or more principals

## Usage

```sh
m365 spo hubsite rights revoke [options]
```

## Options

`-u, --hubSiteUrl <hubSiteUrl>`
: The URL of the hub site to revoke rights on

`-p, --principals <principals>`
: Comma-separated list of principals to revoke join rights. Principals can be users or mail-enabled security groups in the form of `alias` or `alias@<domain name>.com`

`--confirm`
: Don't prompt for confirming revoking rights

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global or SharePoint administrator.

## Examples

Revoke rights to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_ from user with alias _PattiF_. Will prompt for confirmation before revoking the rights

```sh
m365 spo hubsite rights revoke --hubSiteUrl https://contoso.sharepoint.com/sites/sales --principals PattiF
```

Revoke rights to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_ from user with aliases _PattiF_ and _AdeleV_ without prompting for confirmation

```sh
m365 spo hubsite rights revoke --hubSiteUrl https://contoso.sharepoint.com/sites/sales --principals "PattiF,AdeleV" --confirm
```

## Response

The command won't return a response on success.

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
