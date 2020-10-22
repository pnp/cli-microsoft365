# aad o365group renew

Renews Microsoft 365 group's expiration

## Usage

```sh
m365 aad o365group renew [options]
```

## Options

`-i, --id <id>`
: The ID of the Microsoft 365 group to renew

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified _id_ doesn't refer to an existing group, you will get a `The remote server returned an error: (404) Not Found.` error.

## Examples

Renew the Microsoft 365 group with id _28beab62-7540-4db1-a23f-29a6018a3848_

```sh
m365 aad o365group renew --id 28beab62-7540-4db1-a23f-29a6018a3848
```