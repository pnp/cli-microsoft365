# spo contenttype get

Retrieves information about the specified list or site content type

## Usage

```sh
m365 spo contenttype get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: Absolute URL of the site where the content type is located

`-l, --listTitle [listTitle]`
: Title of the list where the content type is located (if it is a list content type)

`-i, --id <id>`
: The ID of the content type to retrieve

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If no content type with the specified is found in the site or the list, you will get the _Content type with ID 0x010012 not found_ error.

## Examples

Retrieve site content type

```sh
m365 spo contenttype get --webUrl https://contoso.sharepoint.com/sites/contoso-sales --id 0x0100558D85B7216F6A489A499DB361E1AE2F
```

Retrieve list content type

```sh
m365 spo contenttype get --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --id 0x0100558D85B7216F6A489A499DB361E1AE2F
```