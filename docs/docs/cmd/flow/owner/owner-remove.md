# flow owner remove

Removes owner permissions to a Power Automate flow

## Usage

```sh
m365 flow owner remove [options]
```

## Options

`-e, --environmentName <environmentName>`
: The name of the environment.

`-n, --name <name>`
: The name of the Power Automate flow.

`--userId [userId]`
: The ID of the user. Specify either `userId`, `userName`, `groupId` or `groupName`.

`--userName [userName]`
: User principal name of the user. Specify either `userId`, `userName`, `groupId` or `groupName`.

`--groupId [groupId]`
: The ID of the group. Specify either `userId`, `userName`, `groupId` or `groupName`.

`--groupName [groupName]`
: The name of the group. Specify either `userId`, `userName`, `groupId` or `groupName`.

`--asAdmin`
: Run the command as admin.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If the environment with the name you specified doesn't exist, you will get the `Access to the environment 'xyz' is denied.` error.

## Examples

Remove owner permissions from a specific Power Automate flow for a user by ID

```sh
m365 flow owner remove --userId "5c241023-2ba5-4ea8-a516-a2481a3e6c51" --environmentName Default-c5a5d746-3520-453f-8a69-780f8e44917e --name 72f2be4a-78c1-4220-a048-dbf557296a72
```

Remove owner permissions from a specific Power Automate flow for a user by UPN

```sh
m365 flow owner remove --userName john.doe@contoso.com --environmentName Default-c5a5d746-3520-453f-8a69-780f8e44917e --name 72f2be4a-78c1-4220-a048-dbf557296a72
```

Remove owner permissions from a specific Power Automate flow for a group by ID

```sh
m365 flow owner remove --groupId "5c241023-2ba5-4ea8-a516-a2481a3e6c51" --environmentName Default-c5a5d746-3520-453f-8a69-780f8e44917e --name 72f2be4a-78c1-4220-a048-dbf557296a72
```

## Response
