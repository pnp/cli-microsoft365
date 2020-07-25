# spo field add

Adds a new list or site column using the CAML field definition

## Usage

```sh
m365 spo field add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: Absolute URL of the site where the field should be created

`-l, --listTitle [listTitle]`
: Title of the list where the field should be created (if it should be created as a list column)

`-x, --xml <xml>`
: CAML field definition

`--options [options]`
: The options to use to add to the field. Allowed values: `DefaultValue`,`AddToDefaultContentType`, `AddToNoContentType`, `AddToAllContentTypes`, `AddFieldInternalNameHint`, `AddFieldToDefaultView`, `AddFieldCheckDisplayName`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified field already exists, you will get a _A duplicate field name "your-field" was found._ error.

## Examples

Create a date time site column

```sh
m365 spo field add --webUrl https://contoso.sharepoint.com/sites/contoso-sales --xml '`<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>`'
```

Create a URL list column

```sh
m365 spo field add --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --xml '`<Field Type="URL" DisplayName="More information link" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Hyperlink" Group="PnP Columns" ID="{6085e32a-339b-4da7-ab6d-c1e013e5ab27}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertMoreInformation" Name="PnPAlertMoreInformation"></Field>`'
```

Create a URL list column and add it to all content types

```sh
m365 spo field add --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --xml '`<Field Type="URL" DisplayName="More information link" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Hyperlink" Group="PnP Columns" ID="{6085e32a-339b-4da7-ab6d-c1e013e5ab27}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertMoreInformation" Name="PnPAlertMoreInformation"></Field>`' --options AddToAllContentTypes
```

## More information

- AddFieldOptions enumeration: [https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.addfieldoptions.aspx](https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.addfieldoptions.aspx)