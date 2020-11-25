# spo file sharinginfo get

Generates a sharing information report for the specified file

## Usage

```sh
m365 spo file sharinginfo get [options]
```

## Options

`-h, --help`
: output usage information

`-w, --webUrl <webUrl>`
: The URL of the site where the file is located

`-u, --url [url]`
: The server-relative URL of the file for which to build the report. Specify either `url` or `id` but not both

`-i, --id [id]`
: The UniqueId (GUID) of the file for which to build the report. Specify either `url` or `id` but not both

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get file sharing information report for the file with server-relative url _/sites/M365CLI/Shared Documents/SharedFile.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file sharinginfo get --webUrl https://contoso.sharepoint.com/sites/project-x --url "/sites/M365CLI/Shared Documents/SharedFile.docx"
```

Get file sharing information report for file with id (UniqueId) _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file sharinginfo get --webUrl https://contoso.sharepoint.com/sites/project-x --id "b2307a39-e878-458b-bc90-03bc578531d6"
```
