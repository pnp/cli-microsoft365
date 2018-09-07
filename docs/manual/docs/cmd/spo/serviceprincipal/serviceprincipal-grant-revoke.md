# spo serviceprincipal grant revoke

Revokes the specified set of permissions granted to the service principal

## Usage

```sh
spo serviceprincipal grant revoke [options]
```

## Alias

```sh
spo sp grant revoke
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --grantId <grantId>`|`ObjectId` of the permission grant to revoke
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To revoke permissions granted to the service principal, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`

The permission grant you want to revoke is denoted using its `ObjectId`. You can retrieve it using the [spo serviceprincipal grant list](./serviceprincipal-grant-list.md) command.

## Examples

Revoke permission grant with ObjectId _50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA_

```sh
spo serviceprincipal grant revoke --grantId 50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA
```