# yammer message get

Returns a Yammer message

## Usage

```sh
m365 yammer message get [options]
```

## Options

`--id <id>`
: The id of the Yammer message

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Returns the Yammer message with the id 1239871123

```sh
m365 yammer message get --id 1239871123
```

Returns the Yammer message with the id 1239871123 in JSON format

```sh
m365 yammer message get --id 1239871123 --output json
```
