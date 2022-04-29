# spo listitem roleinheritance break

Break inheritance of list item. Keeping existing permissions is the default behavior.

## Usage

```sh
m365 spo listitem roleinheritance break [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site from which the item should be retrieved

`-i, --id <id>`
: ID of the item to retrieve

`-l, --listId [listId]`
: ID of the list. Specify listId or listTitle but not both

`-t, --listTitle [listTitle]`
Title of the list. Specify listId or listTitle but not both

`-c, --clearExistingPermissions`
: Flag if used clears all roles from the listitem

--8<-- "docs/cmd/_global.md"

## Examples

Break inheritance of list item _1_ in list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "_someList_" --id 1
```

Break inheritance of list item _1_ in list with ID _202b8199-b9de-43fd-9737-7f213f51c991_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem roleinheritance break --webUrl https://contoso.sharepoint.com/sites/project-x --listId 202b8199-b9de-43fd-9737-7f213f51c991 --id 1
```

Break inheritance of list item _1_ in list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_ with clearing permissions 

```sh
m365 spo listitem roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "_someList_" --id 1 --clearExistingPermissions
```

Break inheritance of list item _1_ in list with ID _202b8199-b9de-43fd-9737-7f213f51c991_ located in site _https://contoso.sharepoint.com/sites/project-x_ with clearing permissions 

```sh
m365 spo listitem roleinheritance break --webUrl https://contoso.sharepoint.com/sites/project-x --listId 202b8199-b9de-43fd-9737-7f213f51c991 --id 1 --clearExistingPermissions
```