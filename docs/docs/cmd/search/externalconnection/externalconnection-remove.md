# search externalconnection remove

Allow the administrator to remove a specific external connection used in Microsoft Search.

## Usage

```sh
m365 search externalconnection remove [options]
```

## Options

`-i, --id [id]`
: ID of the External Connection to remove. Specify either `id` or `name`

`-n, --name [name]`
: Name of the External Connection to remove. Specify either `id` or `name`

`--confirm`
: Don't prompt for confirming removing the connection

--8<-- "docs/cmd/_global.md"

## Remarks

If the command finds multiple external connections used in Microsoft Search with the specified name, it will prompt you to disambiguate which external connection it should remove, listing the discovered IDs.

## Examples

Removes external connection with id _MyApp_

```sh
m365 search externalconnection remove --id "MyApp"
```

Removes external connection with name _Test_. Will NOT prompt for confirmation before removing.

```sh
m365 search externalconnection remove --name "Test" --confirm
```