# logout

Log out from Microsoft 365

## Usage

```sh
m365 logout [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

The `logout` command logs out from Microsoft 365 and removes any access and refresh tokens from memory

## Examples

Log out from Microsoft 365

```sh
m365 logout
```

Log out from Microsoft 365 in debug mode including detailed debug information in the console output

```sh
m365 logout --debug
```