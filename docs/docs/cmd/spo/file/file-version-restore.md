# spo file version restore

Restores a specific version of a specified file

## Usage

```sh
m365 spo file version restore [options]
```

## Options

`-w, --webUrl <webUrl>`
: The URL of the site where the file is located

`--label <label>`
: Label of version which will be restored

`-u, --fileUrl [fileUrl]`
: The server-relative URL of the file whose version will be restored. Specify either `fileUrl` or `fileId` but not both

`-i, --fileId [fileId]`
: The UniqueId (GUID) of the file whose version will be restored. Specify either `fileUrl` or `fileId` but not both

`--confirm [confirm]`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Restores a file version in a specific site based on fileId and prompts for confirmation

```sh
m365 spo file version restore --webUrl https://contoso.sharepoint.com --label "1.0" --fileId 'b2307a39-e878-458b-bc90-03bc578531d6'
```

Restores a file version in a specific site based on fileUrl and prompts for confirmation

```sh
m365 spo file version restore --webUrl https://contoso.sharepoint.com --label "1.0" --fileUrl '/Shared Documents/Document.docx'
```

Restores a file version in a specific site based on fileId without prompting for confirmation

```sh
m365 spo file version restore --webUrl https://contoso.sharepoint.com --label "1.0" --fileId 'b2307a39-e878-458b-bc90-03bc578531d6' --confirm
```

Restores a file version in a specific site based on fileUrl without prompting for confirmation

```sh
m365 spo file version restore --webUrl https://contoso.sharepoint.com --label "1.0" --fileUrl '/Shared Documents/Document.docx' --confirm
```

## Response

The command won't return a response on success.
