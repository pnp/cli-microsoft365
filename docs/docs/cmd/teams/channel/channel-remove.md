# teams channel remove

Removes the specified channel in the Microsoft Teams team

## Usage

```sh
m365 teams channel remove [options]
```

## Options

`-c, --channelId [channelId]`
: The ID of the channel to remove. Specify either `channelId` or `channelName` but not both

`-n, --channelName [channelName]`
: The name of the channel to remove. Specify either `channelId` or `channelName` but not both

`-i, --teamId [teamId]`
: The ID of the team to which the channel to remove belongs. Specify either `teamId` or `teamName` but not both

`--teamName [teamName]`
: The display name of the team to which the channel to remove belongs to. Specify either `teamId` or `teamName` but not both

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Remarks

When deleted, Microsoft Teams channels are moved to a recycle bin and can be restored within 30 days. After that time, they are permanently deleted.

## Examples

Remove the specified Microsoft Teams channel by Id _19:f3dcbb1674574677abcae89cb626f1e6@thread.skype_ from the Microsoft Teams team with id _d66b8110-fcad-49e8-8159-0d488ddb7656_

```sh
m365 teams channel remove --channelId 19:f3dcbb1674574677abcae89cb626f1e6@thread.skype --teamId d66b8110-fcad-49e8-8159-0d488ddb7656
```

Remove the specified Microsoft Teams channel by Id _19:f3dcbb1674574677abcae89cb626f1e6@thread.skype_ from the Microsoft Teams team with id _d66b8110-fcad-49e8-8159-0d488ddb7656_ without confirmation

```sh
m365 teams channel remove --channelId 19:f3dcbb1674574677abcae89cb626f1e6@thread.skype --teamId d66b8110-fcad-49e8-8159-0d488ddb7656 --confirm
```

Remove the specified Microsoft Teams channel by Name _channelName_ from the Microsoft Teams team with name _Team Name_

```sh
m365 teams channel remove --channelName 'channelName' --teamName "Team Name"
```

Remove the specified Microsoft Teams channel by Name _channelName_ from the Microsoft Teams team with name _Team Name_ without confirmation

```sh
m365 teams channel remove --channelName 'channelName' --teamName "Team Name" --confirm 
```

## More information

- directory resource type (deleted items): [https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0)
