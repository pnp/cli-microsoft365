# spo userprofile set

Sets user profile property for a SharePoint user

## Usage

```sh
spo userprofile set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --userName <userName>`|Account name of the user
`-n, --propertyName <propertyName>`|The property name of the property to be set
`-v, --propertyValue <propertyValue>`|The value of the property to be set
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

This command requires tenant admin permissions in case of updating properties other than the current logged user.

## Examples

Updates single value property of a user profile with property name *SPS-JobTitle* and property value 'Senior Developer'

```sh
spo userprofile set --userName 'john.doe@mytenant.onmicrosoft.com' --propertyName 'SPS-JobTitle' --propertyValue 'Senior Developer'
```

Updates multi value property of a user profile with property name *SPS-Skills* and property values 'CSS', 'HTML'

```sh
spo userprofile set --userName 'john.doe@mytenant.onmicrosoft.com' --propertyName 'SPS-Skills' --propertyValue 'CSS, HTML'
```