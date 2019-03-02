# spo listitem record undeclare 

Undeclares  listitem  as a record

## Usage

```sh
spo listitem record undeclare [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the list item should be undeclared as a record
`-i, --id <id>`|ID of the list item to be undeclared as a record.
`-l, --listId [listId]`|ID of the list where the list item should be undeclared as a record. Specify listId or listTitle but not both
`-t, --listTitle [listTitle]`|Title of the list where the list item should be undeclared as a record. Specify listId or listTitle but not both
`-o, --output [output] `|Output type. json|text. Default text
`--verbose `|Runs command with verbose logging
`--debug `| Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the spo login command.
  
## Remarks
  
To undeclare an item as a record in a list, you have to first log in to SharePoint using the spo login command,
eg. o365$ spo login https://contoso.sharepoint.com.
        
## Examples

Undeclare the list item as a record with ID 1 from list with ID  0cd891ef-afce-4e55-b836-fce03286cccf located in site https://contoso.sharepoint.com/sites/project-x 

```sh
o365$ spo listitem record undeclare --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id 1
```

Undeclare the list item as a record with ID 1 from list with title  List 1 located in site https://contoso.sharepoint.com/sites/project-x 

```sh
o365$ spo listitem record undeclare --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --id 1
```