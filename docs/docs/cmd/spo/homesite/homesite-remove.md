# spo homesite remove

Removes the current Home Site

## Usage

```sh
m365 spo homesite remove [options]
```

## Options

`-h, --help`
: output usage information

`--confirm`
: Do not prompt for confirmation before removing the Home Site

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

Removes the current Home Site without confirmation

```sh
m365 spo homesite remove --confirm
```

## More information

- SharePoint home sites, a landing for your organization on the intelligent intranet: [https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933](https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933)
