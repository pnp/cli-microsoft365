# outlook message move

Moves message to the specified folder

## Usage

```sh
m365 outlook message move [options]
```

## Options

`--id <id>`
: ID of the message to move

`--sourceFolderName [sourceFolderName]`
: Name of the folder to move the message from

`--sourceFolderId [sourceFolderId]`
: ID of the folder to move the message from

`--targetFolderName [targetFolderName]`
: Name of the folder to move the message to

`--targetFolderId [targetFolderId]`
: ID of the folder to move the message to

--8<-- "docs/cmd/_global.md"

## Examples

Move the specified message to another folder specified by ID

```sh
m365 outlook message move --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALdyzhHAAA= --sourceFolderId AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEKAAA= --targetFolderId AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAeUO-fAAA=
```

Move the specified message to another folder specified by name

```sh
m365 outlook message move --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALdyzhHAAA= --sourceFolderName Inbox --targetFolderName "Project X"
```

Move the specified message to another folder specified by its well-known
name

```sh
m365 outlook message move --id AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALdyzhHAAA= --sourceFolderName inbox --targetFolderName archive
```

## More information

- Well-known folder names: [https://docs.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0)
