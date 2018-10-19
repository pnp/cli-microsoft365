# spo file add

Uploads file to the specified folder

## Usage

```sh
spo file add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-w, --webUrl <webUrl>`|The URL of the site where the file should be uploaded to
`-f, --folder <folder>`|Site-relative or Server-relative URL to the folder where the file should be uploaded
`-p, --path <path>`|Local path to the file to upload
`-c, --contentType [contentType]`|Content type name or ID to assign to the file
`--checkOut [checkOut]`|If versioning is enabled, this will check out the file first if it exists, upload the file, then check it in again
`--checkInComment [checkInComment]`|Comment to set when checking the file in
`--approve [approve]`|Will automatically approve the uploaded file
`--approveComment [approveComment]`|Comment to set when approving the file
`--publish [publish]`|Will automatically publish the uploaded file
`--publishComment [publishComment]`|Comment to set when publishing the file
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To add a file, you have to first connect to SharePoint using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

This command allows using unknown properties. Each property corresponds to the list item field that should be set when uploading the file. Command properties should be excluded from the SharePoint request.

## Examples

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents'

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in sub folder 'Shared Documents/Sub Folder 1'

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents/Sub Folder 1' --path 'C:\MS365.jpg'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' specifying server relative url for the --folder option

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder '/sites/project-x/Shared Documents' --path 'C:\MS365.jpg'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' with specified content type

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --contentType 'Picture'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents')}, but checks out existing file before the upload

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --checkOut --checkInComment 'check in comment x'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and approves it (when list moderation is enabled)

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --approve --approveComment 'approve comment x'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and publishes it

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --publish --publishComment 'publish comment x'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes single text field value of the list item

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --Title "New Title"
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes person/gorup field and DateTime field values

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --Editor "[{'Key':'i:0#.f|membership|john.smith@contoso.com'}]" --Modified '6/23/2018 10:15 PM'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes hyperlink or picture field

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --URL 'https://contoso.com, Contoso'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes taxonomy field

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --Topic "HR services|c17baaeb-67cd-4378-9389-9d97a945c701"
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes taxonomy multi value field

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --Topic "HR services|c17baaeb-67cd-4378-9389-9d97a945c701;Inclusion ＆ Diversity|66a67671-ed89-44a7-9be4-e80c06b41f35"
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes choice field and multy choice field

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --ChoiceField1 'Option3' --MultyChoiceField1 'Option2;#Option3'
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes person/gorup field that allows multi user selection

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --AllowedUsers "[{'Key':'i:0#.f|membership|john.smith@contoso.com'},{'Key':'i:0#.f|membership|velin.georgiev@contoso.com'}]"
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes yes/no field

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --HasCar true
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes number field and currency field

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --NumberField 100 --CurrencyField 20
```

Adds file MS365.jpg to site 'https://contoso.sharepoint.com/sites/project-x' in folder 'Shared Documents' and changes lookup field and multilookup field

```sh
spo file add --webUrl https://contoso.sharepoint.com/sites/project-x --folder 'Shared Documents' --path 'C:\MS365.jpg' --LookupField 1 --MultiLookupField "2;#;#3;#;#4;#"
```
      
## More information

- Reference to the PnP add file cmdlet:
[https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/add-pnpfile?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/add-pnpfile?view=sharepoint-ps)
        
- Update file metadata with REST API using ValidateUpdateListItem method:
[https://robertschouten.com/2018/04/30/update-file-metadata-with-rest-api-using-validateupdatelistitem-method/](https://robertschouten.com/2018/04/30/update-file-metadata-with-rest-api-using-validateupdatelistitem-method/)        

- List Items System Update options in SharePoint Online:
[https://www.linkedin.com/pulse/list-items-system-update-options-sharepoint-online-andrew-koltyakov/](https://www.linkedin.com/pulse/list-items-system-update-options-sharepoint-online-andrew-koltyakov/)
        