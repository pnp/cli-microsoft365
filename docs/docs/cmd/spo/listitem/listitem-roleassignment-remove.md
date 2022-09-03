# spo listitem roleassignment remove

Removes a role assignment from a listitem.

## Usage

```sh
m365 spo listitem roleassignment remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the listitem is located

`-i, --listId [listId]`
: ID of the list. Specify either listId, listTitle or listUrl but not multiple.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either listId, listTitle or listUrl but not multiple.

`--listUrl [listUrl]`
: Relative URL of the list. Specify either listId, listTitle or listUrl but not multiple.

`--listItemId <listItemId>`
: Id of the listitem to remove the role from.

`--principalId [principalId]`
: SharePoint ID of principal it may be either user id or group id we want to remove permissions Specify principalId only when upn or groupName are not used.

`--upn [upn]`
: upn/email of user. Specify either upn or princpialId.

`--groupName [groupName]`
: enter group name of Azure AD or SharePoint group. Specify either groupName or princpialId.

`--confirm`
: Don't prompt for confirming removing the role assignment

--8<-- "docs/cmd/_global.md"

## Examples

Remove roleassignment from listitem getting list by title based on group name

```sh
m365 spo listitem roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --listTitle "someList" --listItemId 1 --groupName "saleGroup"
```

Remove roleassignment from listitem getting list by title based on principal Id

```sh
m365 spo listitem roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --listTitle "Events" --listItemId 1 --principalId 2
```

Remove roleassignment from listitem getting list by url based on principal Id

```sh
m365 spo listitem roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --listUrl '/sites/contoso-sales/lists/Events' --listItemId 1 --principalId 2
```


Remove roleassignment from listitem getting list by url based on principal Id without prompting for confirmation

```sh
m365 spo listitem roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --listUrl '/sites/contoso-sales/lists/Events' --listItemId 1 --principalId 2 --confirm
```