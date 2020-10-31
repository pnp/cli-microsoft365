# tenant id get

Gets Microsoft 365 tenant ID for the specified domain

## Usage

```sh
m365 tenant id get [options]
```

## Options

`-h, --help`
: output usage information

`-d, --domainName [domainName]`
: The domain name for which to retrieve the Microsoft 365 tenant ID

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If no domain name is specified, the command will return the tenant ID of the tenant to which you are currently logged in.

## Examples

Get Microsoft 365 tenant ID for the specified domain

```sh
m365 tenant id get --domainName contoso.com
```

Get Microsoft 365 tenant ID of the the tenant to which you are currently logged in

```sh
m365 tenant id get
```
