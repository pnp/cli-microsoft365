# spo search

Execute a search query

## Usage

```sh
spo search [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-q, --query <query>`|Query to execute
`-p, --selectProperties`|Comma separated list of properties to retrieve. Will retrieve all properties if not specified and json output is requested.
`--allResults`|Get all results of the search query, not only the amount specified by the rowlimit (default: 10)
`--rowLimit [rowLimit]`|Sets the number of rows to be returned. When the \'allResults\' parameter is enabled, it will determines the size of the batches being retrieved
`--sourceId [sourceId]`|Specifies the identifier (GUID) of the result source to be used to run the query.
`--trimDuplicates`|Specifies whether near duplicate items should be removed from the search results.
`--enableStemming`|Specifies whether stemming is enabled.
`--culture [culture]`|The locale for the query.
`--refinementFilters [refinementFilters]`|The set of refinement filters used when issuing a refinement query. For GET requests, the RefinementFilters parameter is specified as an FQL filter. For POST requests, the RefinementFilters parameter is specified as an array of FQL filters.
`--queryTemplate [queryTemplate]`|A string that contains the text that replaces the query text, as part of a query transformation.
`--sortList [sortList]`|The list of properties by which the search results are ordered.
`--rankingModelId [rankingModelId]`|The ID of the ranking model to use for the query.
`--startRow [startRow]`|The first row that is included in the search results that are returned. You use this parameter when you want to implement paging for search results.
`--properties [properties]`|Additional properties for the query. GET requests support only string values. POST requests support values of any type.
`--sourceName [sourceName]`|Specified the name of the result source to be used to run the query.
`--refiners [refiners]`|The set of refiners to return in a search result.
`-u, --webUrl [webUrl]`|The web against which we want to execute the query. If the parameter is not defined, the query is executed against the web that\'s used when logging in to the SPO environment.
`--hiddenConstraints [hiddenConstraints]`|The additional query terms to append to the query.
`--clientType [clientType]`|The type of the client that issued the query.
`--enablePhonetic`|A Boolean value that specifies whether the phonetic forms of the query terms are used to find matches. (Default = false).
`--processBestBets`|A Boolean value that specifies whether to return best bet results for the query. (Default = false).
`--enableQueryRules`|A Boolean value that specifies whether to enable query rules for the query. (Default = true).
`--processPersonalFavorites`|A Boolean value that specifies whether to return personal favorites with the search results.
`--rawOutput`|Set, to return the unparsed, raw results of the REST call to the search api.
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To execute a search query, you have to first log in to a SharePoint Online site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Execute search query to retrieve all Document Sets (ContentTypeId = '0x0120D520') for the english locale

```sh
spo search --query 'ContentTypeId:0x0120D520' --culture 1033
```

Retrieve all documents. For each document, retrieve the Path, Author and FileType.

```sh
spo search --query 'IsDocument:1' --selectProperties 'Path,Author,FileType' --allResults
```

Return the top 50 items of which the title starts with 'Marketing' while trimming duplicates.

```sh
spo search --query 'Title:Marketing*' --rowLimit=50 --trimDuplicates
```

Return only items from a specific resultsource (using the source id).

```sh
spo search --query '*' --sourceId '6e71030e-5e16-4406-9bff-9c1829843083'
```