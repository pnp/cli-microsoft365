# spo file copy

Copies a file to another location

## Usage

```sh
m365 spo file copy [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located.

`-s, --sourceUrl <sourceUrl>`
: Server-relative or absolute URL of the file.

`-t, --targetUrl <targetUrl>`
: Server-relative or absolute URL of the location.

`--newName [newName]`
: New name of the destination file.

`--nameConflictBehavior [nameConflictBehavior]`
: Behavior when a document with the same name is already present at the destination. Possible values: `fail`, `replace`, `rename`. Default is `fail`.

`--bypassSharedLock`
: This indicates whether a file with a share lock can still be copied. Use this option to copy a file that is locked.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    The required URLs `webUrl`, `sourceUrl` and `targetUrl` cannot be encoded. When you do so, you will get a `File Not Found` error.

When you specify a value for `nameConflictBehavior`, consider the following:

- `fail` will throw an error when the destination file already exists.
- `replace` will replace the destination file if it already exists.
- `rename` will add a suffix (e.g. Document1.pdf) when the destination file already exists.

## Examples

Copy a file to a document library in another site collection with server relative URLs

```sh
m365 spo file copy --webUrl https://contoso.sharepoint.com/sites/project --sourceUrl "/sites/project/Shared Documents/Document.pdf" --targetUrl "/sites/IT/Shared Documents"
```

Copy a file to a document library in another site collection with absolute URLs

```sh
m365 spo file copy --webUrl https://contoso.sharepoint.com/sites/project --sourceUrl "https://contoso.sharepoint.com/sites/project/Shared Documents/Document.pdf" --targetUrl "https://contoso.sharepoint.com/sites/IT/Shared Documents"
```

Copy file to a document library in another site collection with a new name

```sh
m365 spo file copy --webUrl https://contoso.sharepoint.com/sites/project --sourceUrl "/sites/project/Shared Documents/Document.pdf" --targetUrl "/sites/IT/Shared Documents" --newName "Report.pdf"
```

Copy file to a document library in another site collection with a new name, rename the file if it already exists

```sh
m365 spo file copy --webUrl https://contoso.sharepoint.com/sites/project --sourceUrl "/sites/project/Shared Documents/Document.pdf" --targetUrl "/sites/IT/Shared Documents" --newName "Report.pdf" --nameConflictBehavior rename
```

## More information

- Copy items from a SharePoint document library: [https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc](https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc)
