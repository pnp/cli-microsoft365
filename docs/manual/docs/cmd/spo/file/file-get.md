# spo file list

Get information about the specified file

## Usage

```sh
spo file get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-w, --webUrl <webUrl>`|The URL of the site where the folder from which to retrieve files is located
`-u, --url [url]`|server- or site-relative URL of the file. Specify either url or id but not both
`-i, --id [id]`|file ID. Specify either url or id but not both
`--asString`|retrieve the contents of the specified file as string
`--asListItem`|retrieve the underlying list item
`--asFile`|save the file to the path specified in the path option
`-f, --fileName [fileName]`|the name of the file including extension. Must be specified when the --asFile option is used
`-p, --path [path]`|path where to save the file. Must be specified when the --asFile option is used
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get a file, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Return file properties for file with id _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6'
```

Return file as string for file with id _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asString
```

Return list item properties for file with id _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asListItem
```

Save file at path _/Users/user/documents_ with filename _SavedAsTest1.docx_ for file with id _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asFile --path /Users/user/documents --fileName SavedAsTest1.docx
```

Return file properties for file with site relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx'
```

Return file as string for file with site relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asString
```

Return list item properties for file with site relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asListItem
```

Save file at path _/Users/user/documents_ with filename _SavedAsTest1.docx_ for file with site relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file get --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asFile --path /Users/user/documents --fileName SavedAsTest1.docx
```