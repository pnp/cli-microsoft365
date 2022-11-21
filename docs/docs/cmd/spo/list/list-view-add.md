# spo list view add

Adds a new view to a SharePoint list

## Usage

```sh
m365 spo list view add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located.

`--listId [listId]`
: ID of the list to which the view should be added. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`--listTitle [listTitle]`
: Title of the list to which the view should be added. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`--listUrl [listUrl]`
: Relative URL of the list to which the view should be added. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`--title <title>`
: Title of the view to be created for the list.

`--fields <fields>`
: Comma-separated list of **case-sensitive** internal names of the fields to add to the view.

`--viewQuery [viewQuery]`
: XML representation of the list query for the underlying view.

`--personal`
: View will be created as personal view, if specified.

`--default`
: View will be set as default view, if specified.

`--paged`
: View supports paging, if specified (recommended to use this).

`--rowLimit [rowLimit]`
: Sets the number of items to display for the view. Default value is 30.

--8<-- "docs/cmd/_global.md"

## Remarks

We recommend using the `paged` option. When specified, the view supports displaying more items page by page (default behavior). When not specified, the `rowLimit` is absolute, and there is no link to see more items.

## Examples

Add a view called _All events_ to a list with specific title.

```sh
m365 spo list view add --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "My List" --title "All events" --fields "FieldName1,FieldName2,Created,Author,Modified,Editor" --paged
```

Add a view as default view with title _All events_ to a list with a specific URL.

```sh
m365 spo list view add --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl "/Lists/MyList" --title "All events" --fields "FieldName1,Created" --paged --default
```

Add a personal view called _All events_ to a list with a specific ID.

```sh
m365 spo list view add --webUrl https://contoso.sharepoint.com/sites/project-x --listId 00000000-0000-0000-0000-000000000000 --title "All events" --fields "FieldName1,Created" --paged --personal
```

Add a view called _All events_ with defined filter and sorting.

```sh
m365 spo list view add --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "My List" --title "All events" --fields "FieldName1" --viewQuery "<OrderBy><FieldRef Name='Created' Ascending='FALSE' /></OrderBy><Where><Eq><FieldRef Name='TextFieldName' /><Value Type='Text'>Field value</Value></Eq></Where>" --paged
```

## Response

=== "JSON"

    ```json
    {
      "Aggregations": null,
      "AggregationsStatus": null,
      "AssociatedContentTypeId": null,
      "BaseViewId": null,
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
      "HtmlSchemaXml": "<View Type=\"HTML\" Url=\"/Lists/Test/All events.aspx\" Personal=\"FALSE\" DisplayName=\"All events\" DefaultView=\"FALSE\" Name=\"{3CD2E934-F482-4D4A-A9B8-A13B49B3D226}\"><ViewFields><FieldRef Name=\"Title\" /></ViewFields><Query><OrderBy><FieldRef Name=\"Created\" Ascending=\"FALSE\" /></OrderBy><Where><Eq><FieldRef Name=\"TextFieldName\" /><Value Type=\"Text\">Field value</Value></Eq></Where></Query><RowLimit Paged=\"TRUE\">30</RowLimit></View>",
      "Id": "3cd2e934-f482-4d4a-a9b8-a13b49b3d226",
      "ImageUrl": null,
      "IncludeRootFolder": false,
      "ViewJoins": null,
      "JSLink": null,
      "ListViewXml": "<View Type=\"HTML\" Url=\"/Lists/Test/All events.aspx\" Personal=\"FALSE\" DisplayName=\"All events\" DefaultView=\"FALSE\" Name=\"{3CD2E934-F482-4D4A-A9B8-A13B49B3D226}\" ><Query><OrderBy><FieldRef Name=\"Created\" Ascending=\"FALSE\" /></OrderBy><Where><Eq><FieldRef Name=\"TextFieldName\" /><Value Type=\"Text\">Field value</Value></Eq></Where></Query><ViewFields><FieldRef Name=\"Title\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><Toolbar Type=\"None\"/></View>",
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
        "DecodedUrl": "/Lists/Test/All events.aspx"
      },
      "ServerRelativeUrl": "/Lists/Test/All events.aspx",
      "StyleId": null,
      "TabularView": true,
      "Threaded": false,
      "Title": "All events",
      "Toolbar": null,
      "ToolbarTemplateName": null,
      "ViewType": "HTML",
      "ViewData": null,
      "ViewType2": null,
      "VisualizationInfo": null
    }
    ```

=== "Text"

    ```text
    Aggregations             : null
    AggregationsStatus       : null
    AssociatedContentTypeId  : null
    BaseViewId               : null
    CalendarViewStyles       : null
    ColumnWidth              : null
    ContentTypeId            : {"StringValue":"0x"}
    CustomFormatter          : null
    CustomOrder              : null
    DefaultView              : false
    DefaultViewForContentType: false
    EditorModified           : false
    Formats                  : null
    GridLayout               : null
    Hidden                   : false
    HtmlSchemaXml            : <View Type="HTML" Url="/Lists/Test/All events1.aspx" Personal="FALSE" DisplayName="All events" DefaultView="FALSE" Name="{F037FE93-4C74-4ACB-B7B0-71BA599F13C1}"><ViewFields><FieldRef Name="Title" /></ViewFields><Query><OrderBy><FieldRef Name="Created" Ascending="FALSE" /></OrderBy><Where><Eq><FieldRef Name="TextFieldName" /><Value Type="Text">Field value</Value></Eq></Where></Query><RowLimit Paged="TRUE">30</RowLimit></View>
    Id                       : f037fe93-4c74-4acb-b7b0-71ba599f13c1
    ImageUrl                 : null
    IncludeRootFolder        : false
    JSLink                   : null
    ListViewXml              : <View Type="HTML" Url="/Lists/Test/All events1.aspx" Personal="FALSE" DisplayName="All events" DefaultView="FALSE" Name="{F037FE93-4C74-4ACB-B7B0-71BA599F13C1}" ><Query><OrderBy><FieldRef Name="Created" Ascending="FALSE" /></OrderBy><Where><Eq><FieldRef Name="TextFieldName" /><Value Type="Text">Field value</Value></Eq></Where></Query><ViewFields><FieldRef Name="Title" /></ViewFields><RowLimit Paged="TRUE">30</RowLimit><Toolbar Type="None"/></View>
    Method                   : null
    MobileDefaultView        : false
    MobileView               : false
    ModerationType           : null
    NewDocumentTemplates     : null
    OrderedView              : false
    Paged                    : true
    PersonalView             : false
    ReadOnlyView             : false
    RequiresClientIntegration: false
    RowLimit                 : 30
    Scope                    : 0
    ServerRelativePath       : {"DecodedUrl":"/Lists/Test/All events1.aspx"}
    ServerRelativeUrl        : /Lists/Test/All events1.aspx
    StyleId                  : null
    TabularView              : true
    Threaded                 : false
    Title                    : All events
    Toolbar                  : null
    ToolbarTemplateName      : null
    ViewData                 : null
    ViewJoins                : null
    ViewProjectedFields      : null
    ViewQuery                : <OrderBy><FieldRef Name="Created" Ascending="FALSE" /></OrderBy><Where><Eq><FieldRef Name="TextFieldName" /><Value Type="Text">Field value</Value></Eq></Where>
    ViewType                 : HTML
    ViewType2                : null
    VisualizationInfo        : null
    ```

=== "CSV"

    ```csv
    Aggregations,AggregationsStatus,AssociatedContentTypeId,BaseViewId,CalendarViewStyles,ColumnWidth,ContentTypeId,CustomFormatter,CustomOrder,DefaultView,DefaultViewForContentType,EditorModified,Formats,GridLayout,Hidden,HtmlSchemaXml,Id,ImageUrl,IncludeRootFolder,ViewJoins,JSLink,ListViewXml,Method,MobileDefaultView,MobileView,ModerationType,NewDocumentTemplates,OrderedView,Paged,PersonalView,ViewProjectedFields,ViewQuery,ReadOnlyView,RequiresClientIntegration,RowLimit,Scope,ServerRelativePath,ServerRelativeUrl,StyleId,TabularView,Threaded,Title,Toolbar,ToolbarTemplateName,ViewType,ViewData,ViewType2,VisualizationInfo
    ,,,,,,"{""StringValue"":""0x""}",,,,,,,,,"<View Type=""HTML"" Url=""/Lists/Test/All events2.aspx"" Personal=""FALSE"" DisplayName=""All events"" DefaultView=""FALSE"" Name=""{0F11C3F1-E174-4A85-93A9-B4AFB7BD41B6}""><ViewFields><FieldRef Name=""Title"" /></ViewFields><Query><OrderBy><FieldRef Name=""Created"" Ascending=""FALSE"" /></OrderBy><Where><Eq><FieldRef Name=""TextFieldName"" /><Value Type=""Text"">Field value</Value></Eq></Where></Query><RowLimit Paged=""TRUE"">30</RowLimit></View>",0f11c3f1-e174-4a85-93a9-b4afb7bd41b6,,,,,"<View Type=""HTML"" Url=""/Lists/Test/All events2.aspx"" Personal=""FALSE"" DisplayName=""All events"" DefaultView=""FALSE"" Name=""{0F11C3F1-E174-4A85-93A9-B4AFB7BD41B6}"" ><Query><OrderBy><FieldRef Name=""Created"" Ascending=""FALSE"" /></OrderBy><Where><Eq><FieldRef Name=""TextFieldName"" /><Value Type=""Text"">Field value</Value></Eq></Where></Query><ViewFields><FieldRef Name=""Title"" /></ViewFields><RowLimit Paged=""TRUE"">30</RowLimit><Toolbar Type=""None""/></View>",,,,,,,1,,,"<OrderBy><FieldRef Name=""Created"" Ascending=""FALSE"" /></OrderBy><Where><Eq><FieldRef Name=""TextFieldName"" /><Value Type=""Text"">Field value</Value></Eq></Where>",,,30,0,"{""DecodedUrl"":""/Lists/Test/All events2.aspx""}",/Lists/Test/All events2.aspx,,1,,All events,,,HTML,,,
    ```
