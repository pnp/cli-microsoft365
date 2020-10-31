# spo file add

Uploads file to the specified folder

## Usage

```sh
m365 spo file add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site where the file should be uploaded to

`-f, --folder <folder>`
: Site-relative or server-relative URL to the folder where the file should be uploaded

`-p, --path <path>`
: Local path to the file to upload

`-c, --contentType [contentType]`
: Content type name or ID to assign to the file

`--checkOut [checkOut]`
: If versioning is enabled, this will check out the file first if it exists, upload the file, then check it in again

`--checkInComment [checkInComment]`
: Comment to set when checking the file in

`--approve [approve]`
: Will automatically approve the uploaded file

`--approveComment [approveComment]`
: Comment to set when approving the file

`--publish [publish]`
: Will automatically publish the uploaded file

`--publishComment [publishComment]`
: Comment to set when publishing the file

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

This command allows using unknown properties. Each property corresponds to the list item field that should be set when uploading the file.

## Examples

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in sub folder _Shared Documents/Sub Folder 1_

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents/Sub Folder 1' --path 'C:\MS365.jpg'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ specifying server-relative folder url

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder '/sites/project-x/Shared Documents' --path 'C:\MS365.jpg'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ with specified content type

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --contentType 'Picture'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_, but checks out existing file before the upload

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --checkOut --checkInComment 'check in comment x'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and approves it (when list moderation is enabled)

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --approve --approveComment 'approve comment x'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and publishes it

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --publish --publishComment 'publish comment x'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes single text field value of the list item

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --Title "New Title"
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes person/group field and DateTime field values

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --Editor "[{'Key':'i:0#.f|membership|john.smith@contoso.com'}]" --Modified '6/23/2018 10:15 PM'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes hyperlink or picture field

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --URL 'https://contoso.com, Contoso'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes taxonomy field

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --Topic "HR services|c17baaeb-67cd-4378-9389-9d97a945c701"
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes taxonomy multi-value field

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --Topic "HR services|c17baaeb-67cd-4378-9389-9d97a945c701;Inclusion ＆ Diversity|66a67671-ed89-44a7-9be4-e80c06b41f35"
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes choice field and multi-choice field

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --ChoiceField1 'Option3' --MultiChoiceField1 'Option2;#Option3'
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes person/group field that allows multi-user selection

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --AllowedUsers "[{'Key':'i:0#.f|membership|john.smith@contoso.com'},{'Key':'i:0#.f|membership|velin.georgiev@contoso.com'}]"
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes yes/no field

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --HasCar true
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes number field and currency field

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --NumberField 100 --CurrencyField 20
```

Adds file _MS365.jpg_ to site _https://contoso.sharepoint.com/sites/project-x_ in folder _Shared Documents_ and changes lookup field and multi-lookup field

```sh
m365 spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --LookupField 1 --MultiLookupField "2;#;#3;#;#4;#"
```

## More information

- Update file metadata with REST API using ValidateUpdateListItem method: [https://robertschouten.com/2018/04/30/update-file-metadata-with-rest-api-using-validateupdatelistitem-method/](https://robertschouten.com/2018/04/30/update-file-metadata-with-rest-api-using-validateupdatelistitem-method/)
- List Items System Update options in SharePoint Online: [https://www.linkedin.com/pulse/list-items-system-update-options-sharepoint-online-andrew-koltyakov/](https://www.linkedin.com/pulse/list-items-system-update-options-sharepoint-online-andrew-koltyakov/)