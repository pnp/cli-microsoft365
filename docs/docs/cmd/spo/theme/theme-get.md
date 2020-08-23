# spo theme get

Gets custom theme information

## Usage

```sh
spo theme get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|The name of the theme to retrieve
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type `json,text` Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Get information about a theme

```sh
spo theme get --name Contoso-Blue
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)