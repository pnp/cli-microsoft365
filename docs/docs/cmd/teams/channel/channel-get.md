# teams channel get

Gets information about the specific Microsoft Teams team channel

## Usage

```sh
m365 teams channel get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the team to which the channel belongs

`-c, --channelId <channelId>`
: The ID of the channel for which to retrieve more information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples
  
Get information about Microsoft Teams team channel with id _19:493665404ebd4a18adb8a980a31b4986@thread.skype_

```sh
m365 teams channel get --teamId '00000000-0000-0000-0000-000000000000' --channelId '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
```