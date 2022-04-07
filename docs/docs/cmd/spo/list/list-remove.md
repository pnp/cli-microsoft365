# spo list remove

Removes the specified list

## Usage

```sh
m365 spo list remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to remove is located

`-i, --id [id]`
: The ID of the list to remove. Specify either `id` or `title` but not both

`-t, --title [title]`
: Title of the list to remove. Specify either `id` or `title` but not both

`--confirm`
: Don't prompt for confirming removing the list

--8<-- "docs/cmd/_global.md"

## Examples

Remove the list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list remove --webUrl https://contoso.sharepoint.com/sites/project-x --id 0cd891ef-afce-4e55-b836-fce03286cccf
```

Remove the list with title _List 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list remove --webUrl https://contoso.sharepoint.com/sites/project-x --title 'List 1'
```
