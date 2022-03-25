# search externalconnection get

Retrieves an external connection for Microsoft Search by Id or Name.

## Usage

```sh
m365 search externalconnection get [options]
```

## Options

`-i, --id [id]`
: Developer-provided unique ID of the connection within the Azure Active Directory tenant

`-n, --name [name]`
: The display name of the connection to be displayed in the Microsoft 365 admin center. Maximum length of 128 characters

--8<-- "docs/cmd/_global.md"

## Remarks

To retrieve the External Connection, either Id or Name should be supplied but not both.

## Examples

Returns an external connection with id of "TestApp"

```sh
m365 search externalconnection get --id "TestApp"
```

Returns an external connection with name of "Test App"

```sh
m365 search externalconnection get --name "Test App"
```
