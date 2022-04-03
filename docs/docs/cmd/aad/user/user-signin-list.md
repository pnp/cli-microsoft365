# aad user signin list

Retrieves the Azure AD user sign-ins for the tenant

## Usage

```sh
m365 aad user signin list [options]
```

## Options

`-n, --userName [userName]`
: Filter the user sign-ins by given User's UPN (user principal name), eg. `johndoe@example.com`. Specify either userName or userId

`--userId [userId]`
: Filter the user sign-ins by given User's Id. Specify either userName or userId

`--appDisplayName [appDisplayName]`
: Filter the user sign-ins by the given application display name. Specify either appDisplayName or appId

`--appId [appId]`
: Filter the user sign-ins by the given application identifier. Specify either appDisplayName or appId

--8<-- "docs/cmd/_global.md"

## Examples

Get all user's sign-ins in your tenant

```sh
m365 aad user signin list
```

Get all user's sign-ins filter by given user's UPN in the tenant

```sh
m365 aad user signin list --userName 'johndoe@example.com'
```

Get all user's sign-ins filter by given user's Id in the tenant

```sh
m365 aad user signin list --userId '11111111-1111-1111-1111-111111111111'
```

Get all user's sign-ins filter by given application display name in the tenant

```sh
m365 aad user signin list --appDisplayName 'Graph explorer'
```

Get all user's sign-ins filter by given application identifier in the tenant

```sh
m365 aad user signin list --appId '00000000-0000-0000-0000-000000000000'
```

Get all user's sign-ins filter by given user's UPN and application display name in the tenant

```sh
m365 aad user signin list --userName 'johndoe@example.com' --appDisplayName 'Graph explorer'
```

Get all user's sign-ins filter by given user's Id and application display name in the tenant

```sh
m365 aad user signin list --userId '11111111-1111-1111-1111-111111111111' --appDisplayName 'Graph explorer'
```

Get all user's sign-ins filter by given user's UPN and application identifier in the tenant

```sh
m365 aad user signin list --userName 'johndoe@example.com' --appId '00000000-0000-0000-0000-000000000000'
```

Get all user's sign-ins filter by given user's Id and application identifier in the tenant

```sh
m365 aad user signin list --userId '11111111-1111-1111-1111-111111111111' --appId '00000000-0000-0000-0000-000000000000'
```