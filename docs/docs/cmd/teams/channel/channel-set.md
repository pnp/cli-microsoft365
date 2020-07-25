# teams channel set

Updates properties of the specified channel in the given Microsoft Teams team

## Usage

```sh
m365 teams channel set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the team where the channel to update is located

`--channelName <channelName>`
: The name of the channel to update

`--newChannelName [newChannelName]`
: The new name of the channel

`--description [description]`
: The description of the channel

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples
  
Set new description and display name for the specified channel in the given Microsoft Teams team

```sh
m365 teams channel set --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects --description "Channel for new projects"
```

Set new display name for the specified channel in the given Microsoft Teams team

```sh
m365 teams channel set --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects
```