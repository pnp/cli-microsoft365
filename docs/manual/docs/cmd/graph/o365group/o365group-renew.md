# graph o365group renew

Renews Office 365 group's expiration

## Usage

```sh
graph o365group renew [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the Office 365 group to renew
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To renew expiration of a Office 365 group, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command.

If the specified _id_ doesn't refer to an existing group, you will get a `The remote server returned an error: (404) Not Found.` error.

## Examples

Renew the Office 365 group with id _28beab62-7540-4db1-a23f-29a6018a3848_

```sh
graph o365group renew --id 28beab62-7540-4db1-a23f-29a6018a3848
```