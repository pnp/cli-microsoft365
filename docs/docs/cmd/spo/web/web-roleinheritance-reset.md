# spo web roleinheritance reset

Restores role inheritance of subsite.

## Usage

```sh
m365 spo web roleinheritance reset [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site

`--confirm`
: Do not prompt for confirmation before resetting role inheritance.

--8<-- "docs/cmd/_global.md"

## Examples

Restore role inheritance of subsite _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo web roleinheritance reset --webUrl https://contoso.sharepoint.com/sites/project-x
```

Restore role inheritance of subsite _https://contoso.sharepoint.com/sites/project-x_ without prompting for confirmation

```sh
m365 spo web roleinheritance reset --webUrl https://contoso.sharepoint.com/sites/project-x --confirm
```
