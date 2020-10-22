# tenant id get

Gets Microsoft 365 tenant ID for the specified domain

## Usage

```sh
m365 tenant id get [options]
```

## Options

`-d, --domainName [domainName]`
: The domain name for which to retrieve the Microsoft 365 tenant ID

--8<-- "docs/cmd/_global.md"

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
