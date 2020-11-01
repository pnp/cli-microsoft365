# teams conversationmember list

Lists members of a private channel in Microsoft Teams in the current tenant

## Usage

```sh
m365 teams conversationmember list [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId [teamId]`
: The ID of the team where the channel is located. Specify either `teamId` or `teamName`, but not both.

`--teamName [teamName]`
: The name of the team where the channel is located. Specify either `teamId` or `teamName`, but not both.

`-c, --channelId [channelId]`
: The ID of the channel for which to list members. Specify either `channelId` or `channelName`, but not both.
      
`--channelName [channelName]`
: The name of the channel for which to list members. Specify either `channelId` or `channelName`, but not both.

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

You can only see the content of private channels. Regular channels are not supported. The underlying API currently does not support pagination, so if there's too many members to fit into one request, you won't get all the members.

## Examples

List all members of a private channel based on their ids

```sh
m365 teams conversationmember list --teamId 47d6625d-a540-4b59-a4ab-19b787e40593 --channelId 19:586a8b9e36c4479bbbd378e439a96df2@thread.skype
```

List all members of a private channel based on their names

```sh
m365 teams conversationmember list --teamName "Human Resources" --channelName "Private Channel"
```