# spo file sharinglink remove

Removes a specific sharing link to a file

## Usage

```sh
m365 spo file sharinglink remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located.

`--fileUrl [fileUrl]`
: The server-relative (decoded) URL of the file. Specify either `fileUrl` or `fileId` but not both.

`--fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both.

`-i, --id <id>`
: The ID of the sharing link.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Removes a specific sharing link from a file by id without prompting for confirmation.

```sh
m365 spo file sharinglink remove --webUrl https://contoso.sharepoint.com/sites/demo --fileId daebb04b-a773-4baa-b1d1-3625418e3234 --id U1BEZW1vIFZpc2l0b3Jz --confirm
```

Removes a specific sharing link from a file by a specified site-relative URL with prompting for confirmation.

```sh
m365 spo file sharinglink remove --webUrl https://contoso.sharepoint.com/sites/demo --fileUrl 'Shared Documents/document.docx' --id U1BEZW1vIFZpc2l0b3Jz
```

Removes a specific sharing link from a file by a specified server-relative URL with prompting for confirmation.

```sh
m365 spo file sharinglink remove --webUrl https://contoso.sharepoint.com/sites/demo --fileUrl '/sites/demo/Shared Documents/document.docx' --id U1BEZW1vIFZpc2l0b3Jz
```

## Response

The command won't return a response on success.
