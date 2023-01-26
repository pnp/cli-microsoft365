# spo file sharinglink clear

Removes sharing links of a file

## Usage

```sh
m365 spo file sharinglink clear [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located.

`--fileUrl [fileUrl]`
: The server-relative (decoded) URL of the file. Specify either `fileUrl` or `fileId` but not both.

`--fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both.

`--scope [scope]`
: Scope of the sharing link. Possible options are: `anonymous`, `users` or `organization`. If not specified, all links will be removed.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Removes all sharing links from a file specified by id without prompting for confirmation

```sh
m365 spo file sharinglink clear --webUrl https://contoso.sharepoint.com/sites/demo --fileId daebb04b-a773-4baa-b1d1-3625418e3234 --confirm
```

Removes sharing links of type anonymous from a file specified by url with prompting for confirmation

```sh
m365 spo file sharinglink clear --webUrl https://contoso.sharepoint.com/sites/demo --fileUrl '/sites/demo/Shared Documents/document.docx' --scope anonymous
```

## Response

The command won't return a response on success.
