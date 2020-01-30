# spo orgassetslibrary add

Promotes an existing library to become an organization assets library

## Usage

```sh
spo orgassetslibrary add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Promotes an existing library to become an organization assets library

```sh
spo orgassetslibrary add -libraryUrl https://contoso.sharepoint.com/assets
```

  Promotes an existing library to become an organization assets library with Thumbnail

```sh
spo orgassetslibrary -libraryUrl https://contoso.sharepoint.com/assets -thumbnailUrl https://contoso.sharepoint.com/assets/logo.png
```
