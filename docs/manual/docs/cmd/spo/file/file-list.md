# spo file list

Gets all files within the specified folder and site

## Usage

```sh
spo file list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the folder from which to retrieve files is located
`-f, --folder <folder>`|The server- or site-relative URL of the folder from which to retrieve files
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To get all files, you have to first log in to a SharePoint Online site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Return all files from folder _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo file list --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents'
```