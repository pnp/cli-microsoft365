# spo file retentionlabel ensure

Apply a retention label to a file

## Usage

```sh
m365 spo file retentionlabel ensure [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the retentionlabel from a file to apply is located

`--fileUrl [fileUrl]`
: The server-relative URL of the file that should be labelled. Specify either `fileUrl` or `fileId` but not both.

`i, --fileId [fileId]`
: The UniqueId (GUID) of the file that should be labelled. Specify either `fileUrl` or `fileId` but not both.

`--name <name>`
: Name of the retention label to apply to the file.

--8<-- "docs/cmd/_global.md"

## Examples

Applies a retention label to a file based on the label name and the fileUrl

```sh
m365 spo file retentionlabel ensure --webUrl https://contoso.sharepoint.com/sites/project-x --fileUrl '/Shared Documents/Document.docx' --name 'Some label'
```

Applies a retention label to a file based on the label name and the fileId

```sh
m365 spo file retentionlabel ensure --webUrl https://contoso.sharepoint.com/sites/project-x --fileId '26541f96-017c-4189-a604-599e083533b8'  --name 'Some label'
```

## Response

The command won't return a response on success.
