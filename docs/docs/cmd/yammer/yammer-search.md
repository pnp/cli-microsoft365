# yammer search

Returns a list of messages, users, topics and groups that match the specified query.

## Usage

```sh
m365 yammer search [options]
```

## Options

`--queryText <queryText>`
: The query for the search

`--show [show]`
: Specifies the type of data to return when using --output text. Allowed values `summary, messages, users, topics, groups`.

`--limit [limit]`
: Limits the results returned for each item category.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

Using the `--show` option in JSON output is not supported. To filter JSON results, use either a JMESPath query or filter the data yourself after retrieving them.

## Examples

Returns search result for the query `community`

```sh
m365 yammer search --queryText "community"
```

Returns groups that match `community`

```sh
m365 yammer search --queryText "community" --show "groups"
```

Returns topics that match `community`

```sh
m365 yammer search --queryText "community" --show "topics"
```

Returns the first 50 users who match the search query `nuborocks.onmicrosoft.com`

```sh
m365 yammer search --queryText "nuborocks.onmicrosoft.com" --show "users" --limit 50
```

Returns all search results for the query `community`. Stops at 1000 results.

```sh
m365 yammer search --queryText "community" --output json
```

Returns the search results for the query `community` and limits the results to the first 50 entries for each result category.

```sh
m365 yammer search --queryText "community" --output json --limit 50
```
