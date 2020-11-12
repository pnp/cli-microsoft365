# yammer search

Returns a list of messages, users, topics and groups that match the user’s search query.

## Usage

```sh
m365 yammer search [options]
```

## Options

`-h, --help`
: output usage information

`-s, --search <search>`
: The query for the search

`--limit [limit]`
: Limits the results returned for each item category. Can only be used with the --output json option.

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

    The command without `--output json` will just return the search summary for the query.

## Examples

Returns the search result summary for the query `community`

```sh
m365 yammer search --search "community"
```

Returns all search results for the query `community`. Stops at 1000 results. 

```sh
m365 yammer search --search "community" --output json
```

Returns the search results for the query `community` and limits the results to the first 50 entries for each result category.

```sh
m365 yammer search --search "community" --output json --limit 50
```