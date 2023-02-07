# planner roster member remove

Removes a member from a Microsoft Planner Roster

## Usage

```sh
m365 planner roster member remove [options]
```

## Options

`--rosterId <rosterId>`
: ID of the Planner Roster.

`--userId [userId]`
: User's Azure AD ID. Specify either `userId` or `userName` but not both.

`--userName [userName]`
: User's UPN (user principal name, e.g. johndoe@example.com). Specify either `userId` or `userName` but not both.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

!!! attention
    The Planner Roster will be deleted when it doesn't have any users remaining in the membership list because the last user removed themselves. Roster, its plan and all contained tasks will be deleted within 30 days of this operation.

## Examples

Remove a Roster member by its Azure AD ID

```sh
m365 planner roster member remove --rosterId tYqYlNd6eECmsNhN_fcq85cAGAnd --userId 126878e5-d8f9-4db2-951d-d25486488d38
```

Remove a Roster member by its UPN

```sh
m365 planner roster member remove --rosterId tYqYlNd6eECmsNhN_fcq85cAGAnd --userName john.doe@contoso.com
```

## Response

The command won't return a response on success.
