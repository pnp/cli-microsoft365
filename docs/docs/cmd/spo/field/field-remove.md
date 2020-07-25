# spo field remove

Removes the specified list- or site column

## Usage

```sh
m365 spo field remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: Absolute URL of the site where the field to remove is located

`-l, --listTitle [listTitle]`
: Title of the list where the field is located. Specify only one of `listTitle`, `listId` or `listUrl`

`--listId [listId]`
: ID of the list where the field is located. Specify only one of `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or web-relative URL of the list where the field is located. Specify only one of `listTitle`, `listId` or `listUrl`

`-i, --id [id]`
: The ID of the field to remove. Specify id, fieldTitle, or group

`-t, --fieldTitle [fieldTitle]`
: The display name (case-sensitive) of the field to remove. Specify id, fieldTitle, or group

`-g, --group [group]`
: Delete all fields from this group (case-sensitive). Specify id, fieldTitle, or group

`--confirm`
: Don't prompt for confirming removing the field

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Remove the site column with the specified ID, located in site _https://contoso.sharepoint.com/sites/contoso-sales_

```sh
m365 spo field remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --id 5ee2dd25-d941-455a-9bdb-7f2c54aed11b
```

Remove the list column with the specified ID, located in site _https://contoso.sharepoint.com/sites/contoso-sales_. Retrieves the list by its title

```sh
m365 spo field remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --id 5ee2dd25-d941-455a-9bdb-7f2c54aed11b
```

Remove the list column with the specified display name, located in site _https://contoso.sharepoint.com/sites/contoso-sales_. Retrieves the list by its url

```sh
m365 spo field remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listUrl 'Lists/Events' --fieldTitle 'Title'
```

Remove all site columns from group _MyGroup_

```sh
m365 spo field remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --group 'MyGroup'
```
