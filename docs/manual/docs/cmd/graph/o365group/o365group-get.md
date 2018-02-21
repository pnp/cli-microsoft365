# graph o365group get

Gets information about the specified Office 365 Group

## Usage

```sh
graph o365group get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the Office 365 Group to retrieve information for
`--includeSiteUrl`|Set to retrieve the site URL for the group
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To get information about a Office 365 Group, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

## Examples

Get information about the Office 365 Group with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
graph o365group get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```

Get information about the Office 365 Group with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_ and also retrieve the URL of the corresponding SharePoint site

```sh
graph o365group get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844 --includeSiteUrl
```