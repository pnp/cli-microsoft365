# outlook message list

Gets all mail messages from the specified folder

## Usage

```sh
m365 outlook message list [options]
```

## Options

`--folderName [folderName]`
: Name of the folder from which to list messages

`--folderId [folderId]`
: ID of the folder from which to list messages

--8<-- "docs/cmd/_global.md"

## Examples

List all messages in the folder with the specified name

```sh
m365 outlook message list --folderName Archive
```

List all messages in the folder with the specified ID

```sh
m365 outlook message list --folderId AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAFNAAA=
```

List all messages in the folder with the specified well-known name

```sh
m365 outlook message list --folderName inbox
```

## More information

- Well-known folder names: [https://docs.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0)
