# spo homesite set

Sets the specified site as the Home Site

## Usage

```sh
m365 spo homesite set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --siteUrl <siteUrl>`
: The URL of the site to set as Home Site

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

## Examples

Set the specified site as the Home Site

```sh
m365 spo homesite set --siteUrl https://contoso.sharepoint.com/sites/comms
```

## More information

- SharePoint home sites: a landing for your organization on the intelligent intranet: [https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933](https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933)