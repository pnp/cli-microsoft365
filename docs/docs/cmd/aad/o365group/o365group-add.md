# aad o365group add

Creates Microsoft 365 Group

## Usage

```sh
m365 aad o365group add [options]
```

## Options

`-n, --displayName <displayName>`
: Display name for the Microsoft 365 Group

`-d, --description <description>`
: Description for the Microsoft 365 Group

`-m, --mailNickname <mailNickname>`
: Name to use in the group e-mail (part before the `@`)

`--owners [owners]`
: Comma-separated list of Microsoft 365 Group owners

`--members [members]`
: Comma-separated list of Microsoft 365 Group members

`--isPrivate [isPrivate]`
: Set to `true` if the Microsoft 365 Group should be private and to `false` if it should be public (default)

`-l, --logoPath [logoPath]`
: Local path to the image file to use as group logo

--8<-- "docs/cmd/_global.md"

## Remarks

When specifying the path to the logo image you can use both relative and absolute paths. Note, that ~ in the path, will not be resolved and will most likely result in an error.
If an invalid user is provided in the comma-separated list or Owners or Members, the command operation will fail and the Micrsoft 365 Group will not be created.

## Examples

Create a public Microsoft 365 Group

```sh
m365 aad o365group add --displayName Finance --description 'This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.' --mailNickname finance
```

Create a private Microsoft 365 Group

```sh
m365 aad o365group add --displayName Finance --description 'This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.' --mailNickname finance --isPrivate true
```

Create a public Microsoft 365 Group and set specified users as its owners

```sh
m365 aad o365group add --displayName Finance --description 'This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.' --mailNickname finance --owners "DebraB@contoso.onmicrosoft.com,DiegoS@contoso.onmicrosoft.com"
```

Create a public Microsoft 365 Group and set specified users as its members

```sh
m365 aad o365group add --displayName Finance --description 'This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.' --mailNickname finance --members "DebraB@contoso.onmicrosoft.com,DiegoS@contoso.onmicrosoft.com"
```

Create a public Microsoft 365 Group and set its logo

```sh
m365 aad o365group add --displayName Finance --description 'This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.' --mailNickname finance --logoPath images/logo.png
```
