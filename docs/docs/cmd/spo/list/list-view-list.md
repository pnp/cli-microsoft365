# spo list view list

Lists views configured on the specified list

## Usage

```sh
m365 spo list view list [options]
```

## Options

 `-u, --webUrl <webUrl>`
: URL of the site where the list is located

 `-i, --listId [listId]`
: ID of the list for which to list configured views. Specify either `listId`, `listTitle`, or `listUrl`.

 `-t, --listTitle [listTitle]`
: Title of the list for which to list configured views. Specify either `listId`, `listTitle`, or `listUrl`.

 `--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listId` , `listTitle` or `listUrl`.

--8<-- "docs/cmd/_global.md"

## Examples

List all views for a list by title

```sh
m365 spo list view list --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle Documents
```

List all views for a list by ID

```sh
m365 spo list view list --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf
```

List all views for a list by URL

```sh
m365 spo list view list --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl '/sites/project-x/lists/Events'
```

## Response

=== "JSON"

    ```json
    [
      {
        "Aggregations": null,
        "AggregationsStatus": null,
        "AssociatedContentTypeId": null,
        "BaseViewId": "1",
        "CalendarViewStyles": null,
        "ColumnWidth": null,
        "ContentTypeId": {
          "StringValue": "0x"
        },
        "CustomFormatter": null,
        "CustomOrder": null,
        "DefaultView": false,
        "DefaultViewForContentType": false,
        "EditorModified": false,
        "Formats": null,
        "GridLayout": null,
        "Hidden": false,
        "HtmlSchemaXml": "<View Name=\"{0F11C3F1-E174-4A85-93A9-B4AFB7BD41B6}\" Type=\"HTML\" DisplayName=\"All events\" Url=\"/Lists/Test/All events2.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=47\"><ViewFields><FieldRef Name=\"Title\" /></ViewFields><Query><OrderBy><FieldRef Name=\"Created\" Ascending=\"FALSE\" /></OrderBy><Where><Eq><FieldRef Name=\"TextFieldName\" /><Value Type=\"Text\">Field value</Value></Eq></Where></Query><RowLimit Paged=\"TRUE\">30</RowLimit><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><Toolbar Type=\"Standard\" /><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /></ParameterBindings></View>",
        "Id": "0f11c3f1-e174-4a85-93a9-b4afb7bd41b6",
        "ImageUrl": "/_layouts/15/images/generic.png?rev=47",
        "IncludeRootFolder": false,
        "ViewJoins": null,
        "JSLink": "clienttemplates.js",
        "ListViewXml": "<View Name=\"{0F11C3F1-E174-4A85-93A9-B4AFB7BD41B6}\" Type=\"HTML\" DisplayName=\"All events\" Url=\"/Lists/Test/All events2.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=47\" ><Query><OrderBy><FieldRef Name=\"Created\" Ascending=\"FALSE\" /></OrderBy><Where><Eq><FieldRef Name=\"TextFieldName\" /><Value Type=\"Text\">Field value</Value></Eq></Where></Query><ViewFields><FieldRef Name=\"Title\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>",
        "Method": null,
        "MobileDefaultView": false,
        "MobileView": false,
        "ModerationType": null,
        "NewDocumentTemplates": null,
        "OrderedView": false,
        "Paged": true,
        "PersonalView": false,
        "ViewProjectedFields": null,
        "ViewQuery": "<OrderBy><FieldRef Name=\"Created\" Ascending=\"FALSE\" /></OrderBy><Where><Eq><FieldRef Name=\"TextFieldName\" /><Value Type=\"Text\">Field value</Value></Eq></Where>",
        "ReadOnlyView": false,
        "RequiresClientIntegration": false,
        "RowLimit": 30,
        "Scope": 0,
        "ServerRelativePath": {
          "DecodedUrl": "/Lists/Test/All events2.aspx"
        },
        "ServerRelativeUrl": "/Lists/Test/All events2.aspx",
        "StyleId": null,
        "TabularView": true,
        "Threaded": false,
        "Title": "All events",
        "Toolbar": "",
        "ToolbarTemplateName": null,
        "ViewType": "HTML",
        "ViewData": null,
        "ViewType2": null,
        "VisualizationInfo": null
      }
    ]
    ```

=== "Text"

    ```text
    Id                                    Title       DefaultView  Hidden  BaseViewId
    ------------------------------------  ----------  -----------  ------  ----------
    3cd2e934-f482-4d4a-a9b8-a13b49b3d226  All events  false        false   1
    ```

=== "CSV"

    ```csv
    Id,Title,DefaultView,Hidden,BaseViewId
    3cd2e934-f482-4d4a-a9b8-a13b49b3d226,All events,,,1
    ```
