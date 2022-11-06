# spo hubsite disconnect

Disconnect a hub site from its parent hub site

## Usage

```sh
m365 spo hubsite disconnect [options]
```

## Options

`-i, --id [id]`
: ID of the hub site. Specify either `id`, `title` or `url` but not multiple.

`-t, --title [title]`
: Title of the hub site. Specify either `id`, `title` or `url` but not multiple.

`-u, --url [url]`
: Absolute or server-relative URL of the hub site. Specify either `id`, `title` or `url` but not multiple.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global or SharePoint administrator.

To disconnect a regular site from a hub site, use command [spo site hubsite disconnect](../site/site-hubsite-disconnect.md).

If the specified id doesn't point to a valid hub site, you will get a ResourceNotFoundException error.

## Examples

Disconnect a specific hub site with id from its parent hub site

```sh
m365 spo hubsite disconnect --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

Disconnect a specific hub site with url from its parent hub site

```sh
m365 spo hubsite disconnect --url https://contoso.sharepoint.com/sites/project-x
```

Disconnect a specific hub site with title from its parent hub site

```sh
m365 spo hubsite disconnect --title "My hub site"
```

## Response

The command won't return a response on success.
