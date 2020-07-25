# spo contenttype field remove

Removes a column from a site- or list content type

## Usage

```sh
m365 spo contenttype field remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: Absolute URL of the site where the content type is located

`-l, --listTitle [listTitle]`
: Title of the list where the content type is located (if it is a list content type)

`-i, --contentTypeId <contentTypeId>`
: The ID of the content type to remove the column from

`-f, --fieldLinkId <fieldLinkId>`
: The ID of the column to remove

`-c, --updateChildContentTypes`
: Update child content types

`--confirm`
: Don't prompt for confirming removal of a column from content type

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Remove column with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_ from content type with ID _0x0100CA0FA0F5DAEF784494B9C6020C3020A6_ from web with URL _https://contoso.sharepoint.com_

```sh
m365 spo contenttype field remove  --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A6" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --confirm
```

Remove column with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_ from content type with ID _0x0100CA0FA0F5DAEF784494B9C6020C3020A6_ from web with URL _https://contoso.sharepoint.com_ updating child content types

```sh
m365 spo contenttype field remove  --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A6" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --updateChildContentTypes
```

Remove fieldLink with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_ from list content type with ID _0x0100CA0FA0F5DAEF784494B9C6020C3020A6_ from web with URL _https://contoso.sharepoint.com_

```sh
m365 spo contenttype field remove  --contentTypeId "0x0100CA0FA0F5DAEF784494B9C6020C3020A60062F089A38C867747942DB2C3FC50FF6A" --fieldLinkId "880d2f46-fccb-43ca-9def-f88e722cef80" --webUrl https://contoso.sharepoint.com --listTitle "Documents"
```