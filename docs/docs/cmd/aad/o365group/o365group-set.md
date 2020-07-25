# aad o365group set

Updates Microsoft 365 Group properties

## Usage

```sh
m365 aad o365group set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: The ID of the Microsoft 365 Group to update

`-n, --displayName [displayName]`
: Display name for the Microsoft 365 Group

`-d, --description [description]`
: Description for the Microsoft 365 Group

`--owners [owners]`
: Comma-separated list of Microsoft 365 Group owners to add

`--members [members]`
: Comma-separated list of Microsoft 365 Group members to add

`--isPrivate [isPrivate]`
: Set to true if the Microsoft 365 Group should be private and to false if it should be public (default)

`-l, --logoPath [logoPath]`
: Local path to the image file to use as group logo

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

When updating group's owners and members, the command will add newly specified users to the previously set owners and members. The previously set users will not be replaced.

When specifying the path to the logo image you can use both relative and absolute paths. Note, that ~ in the path, will not be resolved and will most likely result in an error.

## Examples

Update Microsoft 365 Group display name

```sh
m365 aad o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --displayName Finance
```

Change Microsoft 365 Group visibility to public

```sh
m365 aad o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --isPrivate false
```

Add new Microsoft 365 Group owners

```sh
m365 aad o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --owners "DebraB@contoso.onmicrosoft.com,DiegoS@contoso.onmicrosoft.com"
```

Add new Microsoft 365 Group members

```sh
m365 aad o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --members "DebraB@contoso.onmicrosoft.com,DiegoS@contoso.onmicrosoft.com"
```

Update Microsoft 365 Group logo

```sh
m365 aad o365group set --id 28beab62-7540-4db1-a23f-29a6018a3848 --logoPath images/logo.png
```
