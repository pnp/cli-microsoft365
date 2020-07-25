# spo site inplacerecordsmanagement set

Activates or deactivates in-place records management for a site collection

## Usage

```sh
m365 spo site inplacerecordsmanagement set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --siteUrl <siteUrl>`
: The URL of the site on which to activate or deactivate in-place records management

`--enabled <enabled>`
: Set to `true` to activate in-place records management and to `false` to deactivate it

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Activates in-place records management for site _https://contoso.sharepoint.com/sites/team-a_

```sh
m365 spo site inplacerecordsmanagement set --siteUrl https://contoso.sharepoint.com/sites/team-a --enabled true
```

Deactivates in-place records management for site _https://contoso.sharepoint.com/sites/team-a_

```sh
m365 spo site inplacerecordsmanagement set --siteUrl https://contoso.sharepoint.com/sites/team-a --enabled false
```