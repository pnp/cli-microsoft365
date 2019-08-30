# teams channel remove

Removes the specified Microsoft Teams channel

## Usage

```sh
teams channel remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-c, --channelId <channelId>`|The ID of the channel to remove
`-i, --teamId <teamId>`|The ID of the team to which the channel to remove belongs
`--confirm`|Don't prompt for confirmation
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

When deleted, Microsoft Teams channels are moved to a recycle bin and can be restored within 30 days. After that time, they are permanently deleted.

## Examples

Removes the specified Teams channel

```sh
teams channel remove --channelId 19:f3dcbb1674574677abcae89cb626f1e6@thread.skype --teamId d66b8110-fcad-49e8-8159-0d488ddb7656
```

Removes the specified Teams channel without confirmation

```sh
teams channel remove --channelId 19:f3dcbb1674574677abcae89cb626f1e6@thread.skype --teamId d66b8110-fcad-49e8-8159-0d488ddb7656 --confirm
```

## More information

- directory resource type (deleted items): [https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0)