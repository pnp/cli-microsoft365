# spo theme set

Add or update a theme

## Usage

```sh
m365 spo theme set [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: Name of the theme to add or update

`-p, --filePath <filePath>`
: Absolute or relative path to the theme json file

`--isInverted`
: Set to specify that the theme is inverted

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

Add or update a theme from a theme JSON file

```sh
m365 spo theme set -n Contoso-Blue -p /Users/rjesh/themes/contoso-blue.json
```

Add or update an inverted theme from a theme JSON file

```sh
m365 spo theme set -n Contoso-Blue -p /Users/rjesh/themes/contoso-blue.json --isInverted
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)
- Theme Generator: [https://aka.ms/themedesigner](https://aka.ms/themedesigner)