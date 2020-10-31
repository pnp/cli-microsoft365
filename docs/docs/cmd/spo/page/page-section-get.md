# spo page section get

Get information about the specified modern page section

## Usage

```sh
m365 spo page section get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located

`-n, --name <name>`
: Name of the page to get section information of

`-s, --section <sectionId>`
: ID of the section for which to retrieve information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified name doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

Get information about the specified section of the modern page named _home.aspx_

```sh
m365 spo page section get --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx --section 1
```