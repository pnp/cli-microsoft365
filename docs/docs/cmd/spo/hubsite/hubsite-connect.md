# spo hubsite connect

Connect a hub site to a parent hub site

## Usage

```sh
m365 spo hubsite connect [options]
```

## Options

`-i, --id [id]`
: ID of the hub site. Specify either `id`, `title`, or `url` but not multiple.

`-t, --title [title]`
: Title of the hub site. Specify either `id`, `title`, or `url` but not multiple.

`-u, --url [url]`
: Absolute or server-relative URL of the hub site. Specify either `id`, `title`, or `url` but not multiple.

`--parentId [parentId]`
: ID of the parent hub site. Specify either `parentId`, `parentTitle`, or `parentUrl` but not multiple.

`--parentTitle [parentTitle]`
: Title of the parent hub site. Specify either `parentId`, `parentTitle`, or `parentUrl` but not multiple.

`--parentUrl [parentUrl]`
: Absolute or server-relative URL of the parent hub site. Specify either `parentId`, `parentTitle`, or `parentUrl` but not multiple.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global or SharePoint administrator.

To connect a regular site to a hub site, use command [spo site hubsite connect](../site/site-hubsite-connect.md).

## Examples

Connect a specific hub site to specific parent hub site by ID.

```sh
m365 spo hubsite connect --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --parentId 637ed2ea-b65b-4a4b-a3d7-ad86953224a4
```

Connect a specific hub site to specific parent hub site by URL.

```sh
m365 spo hubsite connect --url https://contoso.sharepoint.com/sites/project-x --parentUrl https://contoso.sharepoint.com/sites/projects
```

Connect a specific hub site with title to a parent hub site with ID.

```sh
m365 spo hubsite connect --title "My hub site" --parentId 637ed2ea-b65b-4a4b-a3d7-ad86953224a4
```

## Response

The command won't return a response on success.
