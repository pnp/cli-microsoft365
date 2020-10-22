# status

Shows Microsoft 365 login status

## Usage

```sh
m365 status [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

If you are logged in to Microsoft 365, the `status` command will show you information about the user or application name used to sign in and the details about the stored refresh and access tokens and their expiration date and time when run in debug mode.

## Examples

Show the information about the current login to the Microsoft 365

```sh
m365 status
```
