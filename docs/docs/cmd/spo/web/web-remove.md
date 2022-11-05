# spo web remove

Delete specified subsite

## Usage

```sh
m365 spo web remove [options]
```

## Options

`-u, --url <url>`
: URL of the subsite to remove

`--confirm`
: Do not prompt for confirmation before deleting the subsite

--8<-- "docs/cmd/_global.md"

## Examples

Delete subsite without prompting for confirmation

```sh
m365 spo web remove --url https://contoso.sharepoint.com/subsite --confirm
```

## Response

The command won't return a response on success.
