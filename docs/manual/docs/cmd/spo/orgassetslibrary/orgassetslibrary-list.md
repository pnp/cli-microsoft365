# spo orgassetslibrary list

List all libraries that are assigned as asset library

## Usage

```sh
spo orgassetslibrary list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--libraryUrl <libraryUrl>`|The URL of the library to promote
`--thumbnailUrl <thumbnailUrl>`|The URL of the thumbnail to render
`--cdnType [cdnType]`|Specifies the CDN type. Public|Private. Default Private
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

List all libraries that are assigned as asset library

```sh
spo orgassetslibrary list
```
