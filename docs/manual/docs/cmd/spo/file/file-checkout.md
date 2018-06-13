# spo file checkout

Checks out specified file

## Usage

```sh
spo file checkout [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-w, --webUrl <webUrl>`|The URL of the site where the file is located
`-u, --url [url]`|The server-relative URL of the file to retrieve. Specify either `URL` or `UniqueId` but not both
`-i, --id [id]`|The UniqueId (GUID) of the file to retrieve. Specify either `URL` or `UniqueId` but not both
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get a file, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Checks out file with UniqueId _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file checkout --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6'
```

Checks out file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file checkout --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx'
```