# search externalconnection get

Allow the administrator to get a specific external connection for use in Microsoft Search.

## Usage

```sh
m365 search externalconnection get [options]
```

## Options

`-i, --id [id]`
: ID of the External Connection to get. Specify either `id` or `name`

`-n, --name [name]`
: Name of the External Connection to get. Specify either `id` or `name`

--8<-- "docs/cmd/_global.md"

## Examples

Get the External Connection by its id

```sh
m365 search externalconnection get --id "MyApp"
```

Get the External Connection by its name

```sh
m365 search externalconnection get --name "Test"
```
