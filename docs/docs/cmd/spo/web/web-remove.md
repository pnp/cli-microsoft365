# spo web remove

Delete specified subsite

## Usage

```sh
m365 spo web remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the subsite to remove

`--confirm`
: Do not prompt for confirmation before deleting the subsite

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Delete subsite without prompting for confirmation

```sh
m365 spo web remove --webUrl https://contoso.sharepoint.com/subsite --confirm
```