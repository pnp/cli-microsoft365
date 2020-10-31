# spo hubsite rights revoke

Revokes rights to join sites to the specified hub site for one or more principals

## Usage

```sh
m365 spo hubsite rights revoke [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: The URL of the hub site to revoke rights on

`-p, --principals <principals>`
: Comma-separated list of principals to revoke join rights. Principals can be users or mail-enabled security groups in the form of `alias` or `alias@<domain name>.com`

`--confirm`
: Don't prompt for confirming revoking rights

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Revoke rights to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_ from user with alias _PattiF_. Will prompt for confirmation before revoking the rights

```sh
m365 spo hubsite rights revoke --url https://contoso.sharepoint.com/sites/sales --principals PattiF
```

Revoke rights to join sites to the hub site with URL _https://contoso.sharepoint.com/sites/sales_ from user with aliases _PattiF_ and _AdeleV_ without prompting for confirmation

```sh
m365 spo hubsite rights revoke --url https://contoso.sharepoint.com/sites/sales --principals "PattiF,AdeleV" --confirm
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
