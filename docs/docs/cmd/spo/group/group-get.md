# spo group get

Gets site group

## Usage

```sh
m365 spo group get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the group is located.

`-i, --id [id]`
: ID of the site group to get. Use either `id`, `name` or `associatedGroup` but not multiple.

`--name [name]`
: Name of the site group to get. Use either `id`, `name` or `associatedGroup` but not multiple.

`--associatedGroup [associatedGroup]`
: Type of the associated group to get. Available values: `Owner`, `Member`, `Visitor`. Use either `id`, `name` or `associatedGroup` but not multiple.

--8<-- "docs/cmd/_global.md"

## Examples

Get group with ID _7_ for web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --id 7
```

Get group with name _Team Site Members_ for web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --name "Team Site Members"
```

Get the associated owner group of a specified site

```sh
m365 spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --associatedGroup Owner
```
