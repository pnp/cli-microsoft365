# spo web remove

Delete specified subsite

## Usage

```sh
m365 spo web remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the subsite to remove

`--confirm`
: Do not prompt for confirmation before deleting the subsite

--8<-- "docs/cmd/_global.md"

## Examples

Delete subsite without prompting for confirmation

```sh
m365 spo web remove --webUrl https://contoso.sharepoint.com/subsite --confirm
```
