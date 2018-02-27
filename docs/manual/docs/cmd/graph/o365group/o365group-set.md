# graph o365group set

Updates Office 365 Group properties

## Usage

```sh
graph o365group set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the Office 365 Group to update
`-n, --displayName [displayName]`|Display name for the Office 365 Group
`-d, --description [description]`|Description for the Office 365 Group
`--owners [owners]`|Comma-separated list of Office 365 Group owners to add
`--members [members]`|Comma-separated list of Office 365 Group members to add
`--isPrivate [isPrivate]`|Set to true if the Office 365 Group should be private and to false if it should be public (default)
`-l, --logoPath [logoPath]`|Local path to the image file to use as group logo
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To update an Office 365 Group, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

When updating group's owners and members, the command will add newly specified users to the previously set owners and members. The previously set users will not be replaced.

When specifying the path to the logo image you can use both relative and absolute paths. Note, that ~ in the path, will not be resolved and will most likely result in an error.

## Examples

Update Office 365 Group display name

```sh
graph o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --displayName Finance
```

Change Office 365 Group visibility to public

```sh
graph o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --isPrivate false
```

Add new Office 365 Group owners

```sh
graph o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --owners DebraB@contoso.onmicrosoft.com,DiegoS@contoso.onmicrosoft.com
```

Add new Office 365 Group members

```sh
graph o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --members DebraB@contoso.onmicrosoft.com,DiegoS@contoso.onmicrosoft.com
```

Update Office 365 Group logo

```sh
graph o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --logoPath images/logo.png
```