# spo list roleinheritance break

Break inheritance on list or library. Keeping existing permissions is the default behavior.

## Usage

```sh
m365 spo list roleinheritance break [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to retrieve is located

`-i, --listId [listId]`
: ID of the list to retrieve information for. Specify either id or title but not both

`-t, --listTitle [listTitle]`
: Title of the list to retrieve information for. Specify either id or title but not both

`-c, --clearExistingPermissions`
: Flag if used clears all roles from the list

--8<-- "docs/cmd/_global.md"

## Examples

Break inheritance of list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList"
```

Break inheritance of list with ID _202b8199-b9de-43fd-9737-7f213f51c991_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --listId "202b8199-b9de-43fd-9737-7f213f51c991"
```

Break inheritance of list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_ with clearing permissions 

```sh
m365 spo list roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList" --clearExistingPermissions
```

Break inheritance of list with ID _202b8199-b9de-43fd-9737-7f213f51c991_ located in site _https://contoso.sharepoint.com/sites/project-x_ with clearing permissions 

```sh
m365 spo list roleinheritance break --webUrl "https://contoso.sharepoint.com/sites/project-x" --listId "202b8199-b9de-43fd-9737-7f213f51c991" --clearExistingPermissions
```