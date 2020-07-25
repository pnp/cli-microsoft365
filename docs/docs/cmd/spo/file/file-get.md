# spo file get

Gets information about the specified file

## Usage

```sh
m365 spo file get [options]
```

## Options

`-h, --help`
: output usage information

`-w, --webUrl <webUrl>`
: The URL of the site where the file is located

`-u, --url [url]`
: The server-relative URL of the file to retrieve. Specify either `url` or `id` but not both

`-i, --id [id]`
: The UniqueId (GUID) of the file to retrieve. Specify either `url` or `id` but not both

`--asString`
: Set to retrieve the contents of the specified file as string

`--asListItem`
: Set to retrieve the underlying list item

`--asFile`
: Set to save the file to the path specified in the path option

`-p, --path [path]`
: The local path where to save the retrieved file. Must be specified when the `--asFile` option is used

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get file properties for file with id (UniqueId) _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6'
```

Get contents of the file with id (UniqueId) _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asString
```

Get list item properties for file with id (UniqueId) _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asListItem
```

Save file with id (UniqueId) _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_ to local file _/Users/user/documents/SavedAsTest1.docx_

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asFile --path /Users/user/documents/SavedAsTest1.docx
```

Return file properties for file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx'
```

Return file as string for file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asString
```

Return list item properties for file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asListItem
```

Save file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_ to local file _/Users/user/documentsSavedAsTest1.docx_

```sh
m365 spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asFile --path /Users/user/documents/SavedAsTest1.docx
```