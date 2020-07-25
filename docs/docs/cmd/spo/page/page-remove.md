# spo page remove

Removes a modern page

## Usage

```sh
m365 spo page remove [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: Name of the page to remove

`-u, --webUrl <webUrl>`
: URL of the site from which the page should be removed

`--confirm`
: Do not prompt for confirmation before removing the page

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If you try to remove a page with that does not exist, you will get a `The file does not exist` error.

If you set the `--confirm` flag, you will not be prompted for confirmation before the page is actually removed.

## Examples

Remove a modern page.

```sh
m365 spo page remove --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team
```

Remove a modern page without a confirmation prompt.

```sh
m365 spo page remove --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --confirm
```