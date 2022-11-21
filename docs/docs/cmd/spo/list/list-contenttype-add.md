# spo list contenttype add

Adds content type to list

## Usage

```sh
m365 spo list contenttype add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located.

`-i, --listId [listId]`
: ID of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`.

`-i, --id <id>`
: ID of the content type to add to the list

--8<-- "docs/cmd/_global.md"

## Examples

Adds a specific existing content type to a list retrieved by id in a specific site.

```sh
m365 spo list contenttype add --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id 0x0120
```

Adds a specific existing content type to a list retrieved by title in a specific site.

```sh
m365 spo list contenttype add --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents --id 0x0120
```

Adds a specific existing content type to a list retrieved by server relative URL in a specific site.

```sh
m365 spo list contenttype add --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'sites/project-x/Documents' --id 0x0120
```

## Response

=== "JSON"

    ```json
    {
      "ClientFormCustomFormatter": "",
      "Description": "Create a new list item.",
      "DisplayFormClientSideComponentId": "",
      "DisplayFormClientSideComponentProperties": "",
      "DisplayFormTarget": 0,
      "DisplayFormTemplateName": "ListForm",
      "DisplayFormUrl": "",
      "DocumentTemplate": "",
      "DocumentTemplateUrl": "",
      "EditFormClientSideComponentId": "",
      "EditFormClientSideComponentProperties": "",
      "EditFormTarget": 0,
      "EditFormTemplateName": "ListForm",
      "EditFormUrl": "",
      "Group": "List Content Types",
      "Hidden": false,
      "Id": {
        "StringValue": "0x01000B1208C5D23DF44B9F1AEE7373DE9D5E"
      },
      "JSLink": "",
      "MobileDisplayFormUrl": "",
      "MobileEditFormUrl": "",
      "MobileNewFormUrl": "",
      "Name": "Item",
      "NewFormClientSideComponentId": null,
      "NewFormClientSideComponentProperties": "",
      "NewFormTarget": 0,
      "NewFormTemplateName": "ListForm",
      "NewFormUrl": "",
      "ReadOnly": false,
      "SchemaXml": "<ContentType ID=\"0x01000B1208C5D23DF44B9F1AEE7373DE9D5E\" Name=\"Item\" Group=\"List Content Types\" Description=\"Create a new list item.\" Version=\"0\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\" FeatureIds=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162};{c94c1702-30a7-454c-be15-5a895223428d}\"><Folder TargetName=\"Item\"/><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" DisplayName=\"Title\" Required=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" FromBaseType=\"TRUE\" ColName=\"nvarchar1\" ShowInNewForm=\"TRUE\" ShowInEditForm=\"TRUE\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",
      "Scope": "/Lists/Test",
      "Sealed": false,
      "StringId": "0x01000B1208C5D23DF44B9F1AEE7373DE9D5E"
    }
    ```

=== "Text"

    ```text
    ClientFormCustomFormatter               :
    Description                             : Create a new list item.
    DisplayFormClientSideComponentId        :
    DisplayFormClientSideComponentProperties:
    DisplayFormTarget                       : 0
    DisplayFormTemplateName                 : ListForm
    DisplayFormUrl                          :
    DocumentTemplate                        :
    DocumentTemplateUrl                     :
    EditFormClientSideComponentId           :
    EditFormClientSideComponentProperties   :
    EditFormTarget                          : 0
    EditFormTemplateName                    : ListForm
    EditFormUrl                             :
    Group                                   : List Content Types
    Hidden                                  : false
    Id                                      : {"StringValue":"0x01006510BD288ADEDD4AB8AC500FA0B356E4"}
    JSLink                                  :
    MobileDisplayFormUrl                    :
    MobileEditFormUrl                       :
    MobileNewFormUrl                        :
    Name                                    : Item
    NewFormClientSideComponentId            : null
    NewFormClientSideComponentProperties    :
    NewFormTarget                           : 0
    NewFormTemplateName                     : ListForm
    NewFormUrl                              :
    ReadOnly                                : false
    SchemaXml                               : <ContentType ID="0x01006510BD288ADEDD4AB8AC500FA0B356E4" Name="Item" Group="List Content Types" Description="Create a new list item." Version="0" FeatureId="{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}" FeatureIds="{695b6570-a48b-4a8e-8ea5-26ea7fc1d162};{c94c1702-30a7-454c-be15-5a895223428d}"><Folder TargetName="Item"/><Fields><Field ID="{c042a256-787d-4a6f-8a8a-cf6ab767f12d}" Type="Computed" DisplayName="Content Type" Name="ContentType" DisplaceOnUpgrade="TRUE" RenderXMLUsingPattern="TRUE" Sortable="FALSE" SourceID="http://schemas.microsoft.com/sharepoint/v3" StaticName="ContentType" Group="_Hidden" PITarget="MicrosoftWindowsSharePointServices" PIAttribute="ContentTypeID" FromBaseType="TRUE"><FieldRefs><FieldRef Name="ContentTypeId"/></FieldRefs><DisplayPattern><MapToContentType><Column Name="ContentTypeId"/></MapToContentType></DisplayPattern></Field><Field ID="{fa564e0f-0c70-4ab9-b863-0177e6ddd247}" Type="Text" Name="Title" DisplayName="Title" Required="TRUE" SourceID="http://schemas.microsoft.com/sharepoint/v3" StaticName="Title" FromBaseType="TRUE" ColName="nvarchar1" ShowInNewForm="TRUE" ShowInEditForm="TRUE"/></Fields><XmlDocuments><XmlDocument NamespaceURI="http://schemas.microsoft.com/sharepoint/v3/contenttype/forms"><FormTemplates xmlns="http://schemas.microsoft.com/sharepoint/v3/contenttype/forms"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>
    Scope                                   : /Lists/Test
    Sealed                                  : false
    StringId                                : 0x01006510BD288ADEDD4AB8AC500FA0B356E4
    ```

=== "CSV"

    ```csv
    ClientFormCustomFormatter,Description,DisplayFormClientSideComponentId,DisplayFormClientSideComponentProperties,DisplayFormTarget,DisplayFormTemplateName,DisplayFormUrl,DocumentTemplate,DocumentTemplateUrl,EditFormClientSideComponentId,EditFormClientSideComponentProperties,EditFormTarget,EditFormTemplateName,EditFormUrl,Group,Hidden,Id,JSLink,MobileDisplayFormUrl,MobileEditFormUrl,MobileNewFormUrl,Name,NewFormClientSideComponentId,NewFormClientSideComponentProperties,NewFormTarget,NewFormTemplateName,NewFormUrl,ReadOnly,SchemaXml,Scope,Sealed,StringId
    ,Create a new list item.,,,0,ListForm,,,,,,0,ListForm,,List Content Types,,"{""StringValue"":""0x01006D9EF01B2D22B3428279F8CF918B5EE0""}",,,,,Item,,,0,ListForm,,,"<ContentType ID=""0x01006D9EF01B2D22B3428279F8CF918B5EE0"" Name=""Item"" Group=""List Content Types"" Description=""Create a new list item."" Version=""0"" FeatureId=""{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}"" FeatureIds=""{695b6570-a48b-4a8e-8ea5-26ea7fc1d162};{c94c1702-30a7-454c-be15-5a895223428d}""><Folder TargetName=""Item""/><Fields><Field ID=""{c042a256-787d-4a6f-8a8a-cf6ab767f12d}"" Type=""Computed"" DisplayName=""Content Type"" Name=""ContentType"" DisplaceOnUpgrade=""TRUE"" RenderXMLUsingPattern=""TRUE"" Sortable=""FALSE"" SourceID=""http://schemas.microsoft.com/sharepoint/v3"" StaticName=""ContentType"" Group=""_Hidden"" PITarget=""MicrosoftWindowsSharePointServices"" PIAttribute=""ContentTypeID"" FromBaseType=""TRUE""><FieldRefs><FieldRef Name=""ContentTypeId""/></FieldRefs><DisplayPattern><MapToContentType><Column Name=""ContentTypeId""/></MapToContentType></DisplayPattern></Field><Field ID=""{fa564e0f-0c70-4ab9-b863-0177e6ddd247}"" Type=""Text"" Name=""Title"" DisplayName=""Title"" Required=""TRUE"" SourceID=""http://schemas.microsoft.com/sharepoint/v3"" StaticName=""Title"" FromBaseType=""TRUE"" ColName=""nvarchar1"" ShowInNewForm=""TRUE"" ShowInEditForm=""TRUE""/></Fields><XmlDocuments><XmlDocument NamespaceURI=""http://schemas.microsoft.com/sharepoint/v3/contenttype/forms""><FormTemplates xmlns=""http://schemas.microsoft.com/sharepoint/v3/contenttype/forms""><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>",/Lists/Test,,0x01006D9EF01B2D22B3428279F8CF918B5EE0
    ```
