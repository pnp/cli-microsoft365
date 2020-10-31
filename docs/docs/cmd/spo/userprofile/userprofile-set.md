# spo userprofile set

Sets user profile property for a SharePoint user

## Usage

```sh
m365 spo userprofile set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --userName <userName>`
: Account name of the user

`-n, --propertyName <propertyName>`
: The name of the property to be set

`-v, --propertyValue <propertyValue>`
: The value of the property to be set

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

You have to have tenant admin permissions in order to use this command to update profile properties of other users.

## Examples

 Updates the single-value _AboutMe_ property

```sh
m365 spo userprofile set --userName 'john.doe@mytenant.onmicrosoft.com' --propertyName 'AboutMe' --propertyValue 'Working as a Microsoft 365 developer'
```

Updates the multi-value _SPS-Skills_ property

```sh
m365 spo userprofile set --userName 'john.doe@mytenant.onmicrosoft.com' --propertyName 'SPS-Skills' --propertyValue 'CSS, HTML'
```
