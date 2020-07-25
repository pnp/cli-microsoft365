# spo search

Executes a search query

## Usage

```sh
m365 spo search [options]
```

## Options

`-h, --help`
: output usage information

`-q, --queryText <queryText>`
: Query to be executed in KQL format

`-p, --selectProperties [selectProperties]`
: Comma-separated list of properties to retrieve. Will retrieve all properties if not specified and json output is requested.

`-u, --webUrl [webUrl]`
: The web against which we want to execute the query. If the parameter is not defined, the query is executed against the web that's used when logging in to the SPO environment.

`--allResults`
: Set, to get all results of the search query, instead of the number specified by the `rowlimit` (default: 10)

`--rowLimit [rowLimit]`
: The number of rows to be returned. When the `allResults` option is used, the specified value will define the size of retrieved batches

`--sourceId [sourceId]`
: The identifier (GUID) of the result source to be used to run the query.

`--trimDuplicates`
: Set, to remove near duplicate items from the search results.

`--enableStemming`
: Set, to enable stemming.

`--culture [culture]`
: The locale for the query.

`--refinementFilters [refinementFilters]`
: The set of refinement filters used when issuing a refinement query.

`--queryTemplate [queryTemplate]`
: A string that contains the text that replaces the query text, as part of a query transformation.

`--sortList [sortList]`
: The list of properties by which the search results are ordered.

`--rankingModelId [rankingModelId]`
: The ID of the ranking model to use for the query.

`--startRow [startRow]`
: The first row that is included in the search results that are returned. You use this parameter when you want to implement paging for search results.

`--properties [properties]`
: Additional properties for the query.

`--sourceName [sourceName]`
: Specified the name of the result source to be used to run the query.

`--refiners [refiners]`
: The set of refiners to return in a search result.

`--hiddenConstraints [hiddenConstraints]`
: The additional query terms to append to the query.

`--clientType [clientType]`
: The type of the client that issued the query.

`--enablePhonetic`
: Set, to use the phonetic forms of the query terms to find matches. (Default = `false`).

`--processBestBets`
: Set, to return best bet results for the query.

`--enableQueryRules`
: Set, to enable query rules for the query.

`--processPersonalFavorites`
: Set, to return personal favorites with the search results.

`--rawOutput`
: Set, to return the unparsed, raw results of the REST call to the search API.

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Execute search query to retrieve all Document Sets (ContentTypeId = _0x0120D520_) for the English locale

```sh
m365 spo search --queryText "ContentTypeId:0x0120D520" --culture 1033
```

Retrieve all documents. For each document, retrieve the _Path_, _Author_ and _FileType_.

```sh
m365 spo search --queryText "IsDocument:1" --selectProperties "Path,Author,FileType" --allResults
```

Return the top 50 items of which the title starts with _Marketing_ while trimming duplicates.

```sh
m365 spo search --queryText "Title:Marketing*" --rowLimit=50 --trimDuplicates
```

Return only items from a specific result source (using the source id).

```sh
m365 spo search --queryText "*" --sourceId "6e71030e-5e16-4406-9bff-9c1829843083"
```
