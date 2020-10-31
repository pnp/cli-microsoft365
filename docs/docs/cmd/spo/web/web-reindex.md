# spo web reindex

Requests reindexing the specified subsite

## Usage

```sh
m365 spo web reindex [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the subsite to reindex

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the subsite to be reindexed is a no-script site, the command will request reindexing all lists from the subsite that haven't been excluded from the search index.

## Examples

Request reindexing the subsite _https://contoso.sharepoint.com/subsite_

```sh
m365 spo web reindex --webUrl https://contoso.sharepoint.com/subsite
```