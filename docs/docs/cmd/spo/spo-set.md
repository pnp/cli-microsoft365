# spo set

Sets the URL of the root SharePoint site collection for use in SPO commands

## Usage

```sh
m365 spo set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: The URL of the root SharePoint site collection to use in SPO commands

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

CLI for Microsoft 365 automatically discovers the URL of the root SharePoint site collection/SharePoint tenant admin site (whichever is needed to run the particular command). In specific cases, like when managing multi-geo Microsoft 365 tenants, it could be desirable to make the CLI manage the specific geography. For such cases, you can use this command to explicitly specify the SPO URL that should be used when executing SPO commands.

## Examples

Set SPO URL to the specified URL

```sh
m365 spo set --url https://contoso.sharepoint.com
```
