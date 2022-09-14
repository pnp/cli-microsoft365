# booking business get

Retrieve the specified Microsoft Bookings business.

## Usage

```sh
m365 booking business get [options]
```

## Options

`-i, --id [id]`
: ID of the business. Specify either `id` or `name` but not both.

`-n, --name [name]`
: Name of the business. Specify either `id` or `name` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve the specified Microsoft Bookings business with id _business@contoso.onmicrosoft.com_.

```sh
m365 booking business get --id 'business@contoso.onmicrosoft.com'
```

Retrieve the specified Microsoft Bookings business with name _business name_.

```sh
m365 booking business get --name 'business name'
```
