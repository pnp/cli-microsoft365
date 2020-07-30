# spo userprofile set

Get SharePoint user profile properties for the specified user

## Usage

```sh
spo userprofile get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --userName <userName>`|Account name of the user
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

You have to have tenant admin permissions in order to use this command to get profile properties of other users.

## Examples

 Get SharePoint user profile for the specified user

```sh
${commands.USERPROFILE_GET} --userName 'john.doe@mytenant.onmicrosoft.com'
```
