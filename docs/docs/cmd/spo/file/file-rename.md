# spo file rename

Renames a file

## Usage

```sh
m365 spo file rename [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the file is located

`-s, --sourceUrl <sourceUrl>`
: Site-relative URL of the file to copy

`-t, --targetFilename <targetFilename>`
: Filename to which source file be renamed

`--force`
: If a file already exists with target file name, it will be moved to the recycle bin. If omitted, the rename operation will be canceled if the file already exists with same name at the location


--8<-- "docs/cmd/_global.md"

## Examples

Rename a file

```sh
m365 spo file rename --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetFilename sp2.pdf


```

Rename a file with force option, if a file exists with same name then file in recycled

```sh
m365 spo file rename --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetFilename sp2.pdf --force
```

## More information

- Rename a file in a SharePoint document library: [https://support.microsoft.com/en-us/office/rename-a-file-folder-or-link-in-a-document-library-bc493c1a-921f-4bc1-a7f6-985ce11bb185](https://support.microsoft.com/en-us/office/rename-a-file-folder-or-link-in-a-document-library-bc493c1a-921f-4bc1-a7f6-985ce11bb185)