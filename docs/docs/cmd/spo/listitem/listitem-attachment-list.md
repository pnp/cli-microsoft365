# spo listitem attachment list

Gets the attachments associated to a list item

## Usage

```sh
m365 spo listitem attachment list [options]
```

## Options

`-u, --webUrl <webUrl>`
URL of the site from which the item should be retrieved

`--listId [listId]`
: ID of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`--listTitle [listTitle]`
: Title of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

`--itemId <itemId>`
ID of the list item to in question

--8<-- "docs/cmd/_global.md"

## Examples

Gets the attachments from list item with itemId _147_ in list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem attachment list --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List" --itemId 147
```

Gets the attachments from list item with itemId _147_ in list with id _0cd891ef-afce-4e55-b836-fce03286cccf_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem attachment list --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --itemId 147
```

Gets the attachments from a specific list item in a specific list obtained by server-relative URL in a specific site

```sh
m365 spo listitem attachment list --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl /sites/project-x/Documents --itemId 147
```

## Response

=== "JSON"

    ```json
    [
      {
        "FileName": "DummyDocument.docx",
        "FileNameAsPath": {
          "DecodedUrl": "DummyDocument.docx"
        },
        "ServerRelativePath": {
          "DecodedUrl": "/Lists/Test/Attachments/236/DummyDocument.docx"
        },
        "ServerRelativeUrl": "/Lists/Test/Attachments/236/DummyDocument.docx"
      }
    ]
    ```

=== "Text"

    ```text
    FileName            ServerRelativeUrl  
    ------------------  ----------------------------------------------
    DummyDocument.docx  /Lists/Test/Attachments/236/DummyDocument.docx
    ```

=== "CSV"

    ```csv
    FileName,ServerRelativeUrl
    DummyDocument.docx,/Lists/Test/Attachments/236/DummyDocument.docx
    ```
