# spo file checkin

Checks in specified file

## Usage

```sh
spo file checkin [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the file is located
`-f, --fileUrl [fileUrl]`|The server-relative URL of the file to retrieve. Specify either `fileUrl` or `id` but not both
`-i, --id [id]`|The UniqueId (GUID) of the file to retrieve. Specify either `fileUrl` or `id` but not both
`-t, --type [type]`|Type of the check in. Available values Minor|Major|Overwrite. Default is Major
`--comment [comment]`|Comment to set when checking the file in. It\'s length must be less than 1024 letters. Default is empty string
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get a file, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Checks in file with UniqueId _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file checkin --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6'
```

Checks in file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file checkin --webUrl https://contoso.sharepoint.com/sites/project-x --fileUrl '/sites/project-x/documents/Test1.docx'
```

Checks in minor version of file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file checkin --webUrl https://contoso.sharepoint.com/sites/project-x --fileUrl '/sites/project-x/documents/Test1.docx' --type Minor
```

Checks in file _/sites/project-x/documents/Test1.docx_ with comment located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file checkin --webUrl https://contoso.sharepoint.com/sites/project-x --fileUrl '/sites/project-x/documents/Test1.docx' --comment 'approved'
```