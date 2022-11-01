# spo web roleinheritance break

Break role inheritance of subsite

## Usage

```sh
m365 spo web roleinheritance break [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site.

`-c, --clearExistingPermissions`
: Clears all existing permissions from the web.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Remarks

By default, when breaking permissions inheritance, the web will retain existing permissions. To remove existing permissions, use the `--clearExistingPermissions` option.

## Examples

Break role inheritance of a web and keep the existing permissions

```sh
m365 spo web roleinheritance break --webUrl https://contoso.sharepoint.com/sites/project-x
```

Break role inheritance of a web and clear the existing permissions

```sh
m365 spo web roleinheritance break --webUrl https://contoso.sharepoint.com/sites/project-x --clearExistingPermissions
```

## Response

The command won't return a response on success.
