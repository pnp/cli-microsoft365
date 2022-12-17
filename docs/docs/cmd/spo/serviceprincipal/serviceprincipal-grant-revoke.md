# spo serviceprincipal grant revoke

Revokes the specified set of permissions granted to the service principal

## Usage

```sh
m365 spo serviceprincipal grant revoke [options]
```

## Alias

```sh
m365 spo sp grant revoke
```

## Options

`-i, --id <id>`
: `ObjectId` of the permission grant to revoke.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global administrator.

The permission grant you want to revoke is denoted using its `ObjectId`. You can retrieve it using the [spo serviceprincipal grant list](./serviceprincipal-grant-list.md) command.

## Examples

Revoke permission grant

```sh
m365 spo serviceprincipal grant revoke --id 50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA
```

## Response

The command won't return a response on success.
