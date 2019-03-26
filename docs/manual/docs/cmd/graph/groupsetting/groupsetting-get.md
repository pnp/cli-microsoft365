# graph groupsetting get

Gets information about the particular group setting

## Usage

```sh
graph groupsetting get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the group setting to retrieve
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get information about the group setting with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
graph groupsetting get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```