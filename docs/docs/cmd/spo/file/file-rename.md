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
: Site-relative URL of the file to rename

`-t, --targetFileName <targetFileName>`
: New file name of the file

`--force`
: If a file already exists with target file name, it will be moved to the recycle bin. If omitted, the rename operation will be canceled if a file already exists with the specified file name

--8<-- "docs/cmd/_global.md"

## Remarks

If you try to rename a file without the `--force` flag and a file with this name already exists, the operation will be cancelled.

## Examples

Renames a file with server-relative URL _/Shared Documents/Test1.docx_ located in site _<https://contoso.sharepoint.com/sites/project-x>_ to _Test2.docx_

```sh
m365 spo file rename --webUrl https://contoso.sharepoint.com/sites/project-x --sourceUrl '/Shared Documents/Test1.docx' --targetFileName 'Test2.docx'
```

Renames a file with server-relative URL _/Shared Documents/Test1.docx_ located in site _<https://contoso.sharepoint.com/sites/project-x>_ to _Test2.docx_. If the file with the target file name already exists, this file will be moved to the recycle bin

```sh
m365 spo file rename --webUrl https://contoso.sharepoint.com/sites/project-x --sourceUrl '/Shared Documents/Test1.docx' --targetFileName 'Test2.docx' --force
```
