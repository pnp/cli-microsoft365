# spo contenttype list

Lists content types from specified site

## Usage

```sh
m365 spo contenttype list [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site for which to list content types

`-c, --category [category]`
: Category name of content types. When defined will return only content types from specified category

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve site content types

```PowerShell
m365 spo contenttype list --webUrl "https://contoso.sharepoint.com/sites/contoso-sales"
```

Retrieve site content types from the 'List Content Types' category

```PowerShell
m365 spo contenttype list --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --category "List Content Types"
```
