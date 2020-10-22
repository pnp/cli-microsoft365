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

`-i, --grantId <grantId>`
: `ObjectId` of the permission grant to revoke

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

The permission grant you want to revoke is denoted using its `ObjectId`. You can retrieve it using the [spo serviceprincipal grant list](./serviceprincipal-grant-list.md) command.

## Examples

Revoke permission grant with ObjectId _50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA_

```sh
m365 spo serviceprincipal grant revoke --grantId 50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA
```