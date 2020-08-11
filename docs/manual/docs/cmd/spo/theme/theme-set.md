# spo theme set

Add or update a theme

## Usage

```sh
spo theme set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the theme to add or update
`-p, --filePath <filePath>`|Absolute or relative path to the theme json file
`--isInverted`|Set to specify that the theme is inverted
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Add or update a theme from a theme JSON file

```sh
spo theme set -n Contoso-Blue -p /Users/rjesh/themes/contoso-blue.json
```

Add or update an inverted theme from a theme JSON file

```sh
spo theme set -n Contoso-Blue -p /Users/rjesh/themes/contoso-blue.json --isInverted
```

A valid theme object is as follows:

```json
{
    "themePrimary": "#d81e05",
    "themeLighterAlt": "#fdf5f4",
    "themeLighter": "#f9d6d2",
    "themeLight": "#f4b4ac",
    "themeTertiary": "#e87060",
    "themeSecondary": "#dd351e",
    "themeDarkAlt": "#c31a04",
    "themeDark": "#a51603",
    "themeDarker": "#791002",
    "neutralLighterAlt": "#eeeeee",
    "neutralLighter": "#f5f5f5",
    "neutralLight": "#e1e1e1",
    "neutralQuaternaryAlt": "#d1d1d1",
    "neutralQuaternary": "#c8c8c8",
    "neutralTertiaryAlt": "#c0c0c0",
    "neutralTertiary": "#c2c2c2",
    "neutralSecondary": "#858585",
    "neutralPrimaryAlt": "#4b4b4b",
    "neutralPrimary": "#333333",
    "neutralDark": "#272727",
    "black": "#1d1d1d",
    "white": "#f5f5f5"
}
```

Validation checks the following:

- The specified string is a valid JSON string
- The deserialized object contains all properties defined in the above example
- The deserialized object doesn't contain any other properties
- Each property of the deserialized object contains a valid hex color value prefixed with a #

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)
- Theme Generator: [https://aka.ms/themedesigner](https://aka.ms/themedesigner)