# spo file version clear

Removes all version history of a specified file

## Usage

```sh
m365 spo file version clear [options]
```

## Options

`-w, --webUrl <webUrl>`
: The URL of the site where the file is located

`-u, --fileUrl [fileUrl]`
: The server-relative URL of the file to retrieve. Specify either `fileUrl` or `fileId` but not both

`-i, --fileId [fileId]`
: The UniqueId (GUID) of the file to retrieve. Specify either `fileUrl` or `fileId` but not both

`--confirm [confirm]`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Removes all file version history in a specific site based on fileId and prompts for confirmation

```sh
m365 spo file version clear --webUrl https://contoso.sharepoint.com --fileId 'b2307a39-e878-458b-bc90-03bc578531d6'
```

Removes all file version history in a specific site based on fileUrl and prompts for confirmation

```sh
m365 spo file version clear --webUrl https://contoso.sharepoint.com --fileUrl '/Shared Documents/Document.docx'
```

Removes all file version history in a specific site based on fileId without prompting for confirmation

```sh
m365 spo file version clear --webUrl https://contoso.sharepoint.com --fileId 'b2307a39-e878-458b-bc90-03bc578531d6' --confirm
```

Removes all file version history in a specific site based on fileUrl without prompting for confirmation

```sh
m365 spo file version clear --webUrl https://contoso.sharepoint.com --fileUrl '/Shared Documents/Document.docx' --confirm
```

## Response

The command won't return a response on success.
