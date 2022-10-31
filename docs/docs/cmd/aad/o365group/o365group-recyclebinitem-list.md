# aad o365group recyclebinitem list

Lists Groups from the recycle bin in the current tenant

## Usage

```sh
m365 aad o365group recyclebinitem list [options]
```

## Options

`-d, --groupDisplayName [groupDisplayName]`
: Lists groups with DisplayName starting with the specified value

`-m, --groupMailNickname [groupMailNickname]`
: Lists groups with MailNickname starting with the specified value

--8<-- "docs/cmd/_global.md"

## Examples

List all deleted Microsoft 365 Groups in the tenant

```sh
m365 aad o365group recyclebinitem list
```

List deleted Microsoft 365 Groups with display name starting with _Project_

```sh
m365 aad o365group recyclebinitem list --groupDisplayName Project
```

List deleted Microsoft 365 Groups mail nick name starting with _team_

```sh
m365 aad o365group recyclebinitem list --groupMailNickname team
```

List deleted Microsoft 365 Groups mail nick name starting with _team_ and with display name starting with _Project_

```sh
m365 aad o365group recyclebinitem list --groupMailNickname team --groupDisplayName Project
```
