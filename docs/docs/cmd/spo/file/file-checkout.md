# spo file checkout

Checks out specified file

## Usage

```sh
m365 spo file checkout [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located

`-f, --fileUrl [fileUrl]`
: The server-relative URL of the file to retrieve. Specify either `fileUrl` or `id` but not both

`-i, --id [id]`
: The UniqueId (GUID) of the file to retrieve. Specify either `fileUrl` or `id` but not both

--8<-- "docs/cmd/_global.md"

## Examples

Checks out file with UniqueId _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file checkout --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6'
```

Checks out file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file checkout --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx'
```