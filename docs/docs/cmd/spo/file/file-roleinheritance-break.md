# spo file roleinheritance break

Breaks inheritance of file. Keeping existing permissions is the default behavior.

## Usage

```sh
m365 spo file roleinheritance break [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the item for which to break role inheritance is located

`--fileUrl [fileUrl]`
: The server-relative URL of the file to retrieve. Specify either url or id but not both

`i, --fileId [fileId]`
: The UniqueId (GUID) of the file to retrieve. Specify either url or id but not both

`-c, --clearExistingPermissions`
: Set to clear existing roles from the list item

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Examples

Break inheritance of file with id (UniqueId) _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --fileId "b2307a39-e878-458b-bc90-03bc578531d6"
```

Break inheritance of file with id (UniqueId) _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/project-x_ with clearing permissions 

```sh
m365 spo file roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --fileId "b2307a39-e878-458b-bc90-03bc578531d6" --clearExistingPermissions
```

Break inheritance of file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo file roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --fileUrl "/sites/project-x/documents/Test1.docx"
```

Break inheritance of file with server-relative url _/sites/project-x/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/project-x_ with clearing permissions 

```sh
m365 spo file roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --fileUrl "/sites/project-x/documents/Test1.docx" --clearExistingPermissions
```
