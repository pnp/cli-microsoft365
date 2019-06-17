# graph planner task list

Lists Planner tasks of the user

## Usage

```sh
graph planner task list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--userId [userId]`| Retrieves all the tasks of the user. Specify `userId` or `userName` but not both. If none of them are specified, current user tasks will be returned.
`--userName  [userName ]`| Retrieves all the tasks of the user. Specify `userId` or `userName` but not both. If none of them are specified, current user tasks will be returned.
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging.
`--debug`|Runs command with debug logging.

!!! important
    before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To list planner tasks of a user, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

Using the `--userId` and `--userName` option, you can retrieve all the planner tasks of the specified user. But it will result in error if you don't have access to view specific user's task.  You can retrieve information about a user's task, either by specifying that user's id or username (`userPrincipalName`), but not both.

Both userId and userName is optional, if no values are passed for those parameters it will list all the tasks of current logged in user.

## Examples

List all the tasks of current logged in user

```sh
graph planner task list
```

List all tasks of the user with id _1caf7dcd-7e83-4c3a-94f7-932a1299c844_

```sh
graph planner task list --userId 1caf7dcd-7e83-4c3a-94f7-932a1299c844
```

List all tasks of the user with user name _AarifS@contoso.onmicrosoft.com_

```sh
graph planner task list --userName AarifS@contoso.onmicrosoft.com
```

## More information

- Microsoft Graph Get Tasks of User: 
[https://docs.microsoft.com/en-us/graph/api/planneruser-list-tasks?view=graph-rest-1.0&tabs=cs](https://docs.microsoft.com/en-us/graph/api/planneruser-list-tasks?view=graph-rest-1.0&tabs=cs)