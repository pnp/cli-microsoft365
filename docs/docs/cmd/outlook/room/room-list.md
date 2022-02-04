# outlook room list

Get a collection of all available rooms

## Usage

```sh
m365 outlook room list [options]
```

## Options

`--roomlistEmail, [roomlistEmail]`
: Use to filter returned rooms by their roomlist email (eg. bldg2@contoso.com)

--8<-- "docs/cmd/_global.md"

## Examples

Get all the rooms

```sh
m365 outlook room list
```

Get all the rooms of specified roomlist e-mail address

```sh
m365 outlook room list --roomlistEmail "bldg2@contoso.com"
```
