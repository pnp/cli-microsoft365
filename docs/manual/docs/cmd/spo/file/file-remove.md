# spo file remove

Removes the specified file

## Usage

```sh
spo file remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the file to remove is located
`-i, --id [id]`|The ID of the file to remove. Specify either `id` or `url` but not both
`-u, --url [url]`|The server or site-relative URL of the file to remove. Specify either `id` or `url` but not both
`--recycle`|Recycle the file
`--confirm`|Don't prompt for confirming removing the file
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To remove a file, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Remove the file with ID 0cd891ef-afce-4e55-b836-fce03286cccf located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file remove --webUrl https://contoso.sharepoint.com/sites/project-x --id 0cd891ef-afce-4e55-b836-fce03286cccf
```

Remove the file with site-relative url _SharedDocuments/Test.docx_ from list with title _List 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file remove --webUrl https://contoso.sharepoint.com/sites/project-x --url SharedDocuments/Test.docx
```

Remove the file with server-relative url _SharedDocuments/Test.docx_ from list with title _List 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file remove --webUrl https://contoso.sharepoint.com/sites/project-x --url /sites/project-x/SharedDocuments/Test.docx
```