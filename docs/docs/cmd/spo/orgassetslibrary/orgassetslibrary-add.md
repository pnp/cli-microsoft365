# spo orgassetslibrary add

Promotes an existing library to become an organization assets library

## Usage

```sh
m365 spo orgassetslibrary add [options]
```

## Options

`--libraryUrl <libraryUrl>`
: The URL of the library to promote.

`--thumbnailUrl <thumbnailUrl>`
: The URL of the thumbnail to render.

`--cdnType [cdnType]`
: Specifies the CDN type. Allowed values `Public`, `Private`. Default `Private`.

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Promotes an existing library to become an organization assets library

```sh
m365 spo orgassetslibrary add --libraryUrl "https://contoso.sharepoint.com/SiteAssets"
```

Promotes an existing library to become an organization assets library with Thumbnail

```sh
m365 spo orgassetslibrary --libraryUrl "https://contoso.sharepoint.com/SiteAssets" --thumbnailUrl "https://contoso.sharepoint.com/assets/logo.png"
```

## Response

The command won't return a response on success.
