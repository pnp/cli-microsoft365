# Use CLI for Microsoft 365 context

CLI for Microsoft 365 context may store any kind of option and its value. 

## How does it work

When a command is executed CLI will first check for the `.m365rc.json` file in the current directory (this is where the context is saved once you set any option). If present, CLI will parse the context object and check if any of the options defined in it may be used for the currently executed command. If yes CLI will execute this command with the option and its value taken from the context. 

For example if we have the following context defined in the `.m365rc.json` file
```json
{
  "context": {
    "groupName": "test group",
    "listTitle": "test list"
  }
}
```

And we will execute the following command: 
```powershell
m365 spo listitem list --webUrl "https://contoso.sharepoint.com/sites/sample"
```

The command is missing the `listTitle` required option but instead of failing it will be executed as this option is present in the current context. The `groupName` option will not be used from the context as the `spo listitem list` command does not have it.

When the same option is defined in the context and also in the command itself then the value defined in the command will be used.

## Guidance

In order to create an empty context we may execute the following command:
```powershell
m365 context init
```

To add or update an option in the context use the `m365 context option set` command. The `name` is used to define the option name and `value` is used to define its default value. For example, if we want to set `test list` value for option `listTitle` we should execute:
```powershell
m365 context option set --name 'listTitle' --value 'test list'
```

In order to check what is defined in the context we may use the `m365 context option list` command.

To remove a specific option defined in the context we may run:
```powershell
m365 context option remove
```

In order to remove the full context we may execute the following command:
```powershell
m365 context remove
```

## Example

Considering the below context:
```json
{
  "context": {
    "groupName": "test group",
    "listTitle": "test list"
  }
}
```

When we execute:
```powershell
m365 context option set --name "webUrl" --value "https://contoso.sharepoint.com/sites/sample"
```

The result of the above will be a new option added to the context:
```json
{
  "context": {
    "groupName": "test group",
    "listTitle": "test list",
    "webUrl": "https://contoso.sharepoint.com/sites/sample"
  }
}
```

Next, if we execute the following:
```powershell
m365 spo listitem list
```

As a result, we will get a list of items from list `test list` from `https://contoso.sharepoint.com/sites/sample` even though we did not specify any of the required options, those were taken from the context. 

Now when we execute:
```powershell
m365 spo listitem list --listTitle "second list"
```

We will get all items from list `second list` from `https://contoso.sharepoint.com/sites/sample` site.

If we execute:
```powershell
m365 spo list get
```

The command will fail. Although the `webUrl` option will be used, which is the required one, but the `listTitle` will not be used as the command needs `title` option instead to get the specified list.

## Related commands

- [m365 context init](../cmd/context/context-init.md)
- [m365 context remove](../cmd/context/context-remove.md)
- [m365 context option set](../cmd/context/option/option-set.md)
- [m365 context option list](../cmd/context/option/option-list.md)
- [m365 context option remove](../cmd/context/option/option-remove.md)
