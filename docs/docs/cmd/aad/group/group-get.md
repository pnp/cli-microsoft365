# aad group get

Gets information about the specified Azure AD Group

## Usage

```sh
m365 aad group get [options]
```

## Options

`-i, --id [id]`
: The object Id of the Azure AD group. Specify either `id` or `title` but not both

`-t, --title [title]`
: The display name of the Azure AD group. Specify either `id` or `title` but not both

--8<-- "docs/cmd/_global.md"

## Examples

Get information about a Azure AD Group by id

```sh
m365 aad group get --id 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```

Get information about a Azure AD Group by title

```sh
m365 aad group get --title "Finance"
```
