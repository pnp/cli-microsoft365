# spo page remove

Removes a modern page

## Usage

```sh
m365 spo page remove [options]
```

## Options

`-n, --name <name>`
: Name of the page to remove

`-u, --webUrl <webUrl>`
: URL of the site from which the page should be removed

`--confirm`
: Do not prompt for confirmation before removing the page

--8<-- "docs/cmd/_global.md"

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