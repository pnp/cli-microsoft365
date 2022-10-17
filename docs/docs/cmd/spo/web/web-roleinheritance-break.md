# spo web roleinheritance break

Break role inheritance of subsite.

## Usage

```sh
m365 spo web roleinheritance break [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site

`-c, --clearExistingPermissions`
: Flag if used clears all roles from the web

## Remarks

By default, when breaking permissions inheritance, the web will retain existing permissions. To remove existing permissions, use the `--clearExistingPermissions` option.

--8<-- "docs/cmd/_global.md"

## Examples

Break role inheritance of subsite _<https://contoso.sharepoint.com/sites/project-x>_ --confirm

```sh
m365 spo web roleinheritance break --webUrl https://contoso.sharepoint.com/sites/project-x  --confirm
```

Break inheritance of web in site _<https://contoso.sharepoint.com/sites/project-x>_ with clearing permissions --confirm

```sh
m365 spo web roleinheritance break --webUrl https://contoso.sharepoint.com/sites/project-x --clearExistingPermissions --confirm 
```
