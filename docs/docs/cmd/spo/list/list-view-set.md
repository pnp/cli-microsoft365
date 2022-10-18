# spo list view set

Updates existing list view

## Usage

```sh
m365 spo list view set [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`--listId [listId]`
: ID of the list where the view is located. Specify `listTitle` or `listId` but not both

`--listTitle [listTitle]`
: Title of the list where the view is located. Specify `listTitle` or `listId` but not both

`--viewId [viewId]`
: ID of the view to update. Specify `viewTitle` or `viewId` but not both

`--viewTitle [viewTitle]`
: Title of the view to update. Specify `viewTitle` or `viewId` but not both

--8<-- "docs/cmd/_global.md"

## Remarks

Specify properties to update using their names, eg. `--Title 'New Title' --JSLink jslink.js`

When updating list formatting, the value of the CustomFormatter property must be XML-escaped, eg. `&lt;` instead of `<`.

!!! warning "Escaping JSON in PowerShell"
    When updating list view formatting for a view with the `--CustomFormatter` option, it's possible to enter a JSON string. In PowerShell 5 to 7.2 [specific escaping rules](./../../user-guide/using-cli.md#escaping-double-quotes-in-powershell) apply due to an issue. Remember that you can also use [file tokens](./../../user-guide/using-cli.md#passing-complex-content-into-cli-options) instead.

## Examples

Update the title of the list view specified by its name

```sh
m365 spo list view set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --viewTitle 'All items' --Title 'All events'
```

Update the title of the list view specified by its ID

```sh
m365 spo list view set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --viewId 330f29c5-5c4c-465f-9f4b-7903020ae1ce --Title 'All events'
```

Update view formatting of the specified list view

=== "PowerShell"

    ```sh
    m365 spo list view set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --viewTitle 'All items' --CustomFormatter '{\"schema\":\"https://developer.microsoft.com/json-schemas/sp/view-formatting.schema.json\",\"additionalRowClass\": \"=if([$DueDate] &lt;= @now, ''sp-field-severity--severeWarning'', '''')\"}'
    ```

=== "Bash"

    ```sh
    m365 spo list view set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List' --viewTitle 'All items' --CustomFormatter "{\"schema\":\"https://developer.microsoft.com/json-schemas/sp/view-formatting.schema.json\",\"additionalRowClass\": \"=if([$DueDate] &lt;= @now, 'sp-field-severity--severeWarning', '')\"}"
    ```
