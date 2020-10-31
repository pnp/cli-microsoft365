# spo orgassetslibrary remove

Removes a library that was designated as a central location for organization assets across the tenant.

## Usage

```sh
m365 spo orgassetslibrary remove [options]
```

## Options

`-h, --help`
: output usage information

`--libraryUrl <libraryUrl>`
: The server relative URL of the library to be removed as a central location for organization assets

`--confirm`
: Don't prompt for confirming removing the organization asset library

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Removes organization assets library without confirmation

```sh
m365 spo orgassetslibrary remove --libraryUrl "/sites/branding/assets" --confirm
```
