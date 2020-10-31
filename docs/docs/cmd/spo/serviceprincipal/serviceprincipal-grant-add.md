# spo serviceprincipal grant add

Grants the service principal permission to the specified API

## Usage

```sh
m365 spo serviceprincipal grant add [options]
```

## Alias

```sh
m365 spo sp grant add
```

## Options

`-h, --help`
: output usage information

`-r, --resource <resource>`
: The name of the resource for which permissions should be granted

`-s, --scope <scope>`
: The name of the permission that should be granted

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

Grant the service principal permission to read email using the Microsoft Graph

```sh
m365 spo serviceprincipal grant add --resource 'Microsoft Graph' --scope 'Mail.Read'
```

Grant the service principal permission to a custom API

```sh
m365 spo serviceprincipal grant add --resource 'contoso-api' --scope 'user_impersonation'
```