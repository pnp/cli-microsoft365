# spo web reindex

Requests reindexing the specified subsite

## Usage

```sh
m365 spo web reindex [options]
```

## Options

`-u, --url <url>`
: URL of the subsite to reindex

--8<-- "docs/cmd/_global.md"

## Remarks

If the subsite to be reindexed is a no-script site, the command will request reindexing all lists from the subsite that haven't been excluded from the search index.

## Examples

Request reindexing the subsite _https://contoso.sharepoint.com/subsite_

```sh
m365 spo web reindex --url https://contoso.sharepoint.com/subsite
```

## Response

The command won't return a response on success.
