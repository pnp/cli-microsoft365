# teams team clone

Creates a clone of a Microsoft Teams team

## Usage

```sh
m365 teams team clone [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the Microsoft Teams team to clone

`-n, --displayName <displayName>`
: The display name for the new Microsoft Teams Team

`-p, --partsToClone <partsToClone>`
: A comma-separated list of the parts to clone. Allowed values are `apps,channels,members,settings,tabs`

`-d, --description [description]`
: The description for the new Microsoft Teams Team

`-c, --classification [classification]`
: The classification for the new Microsoft Teams Team. If not specified, will be copied from the original Microsoft Teams Team

`-v, --visibility [visibility]`
: Specify the visibility of the new Microsoft Teams Team. Allowed values are `Private,Public`.

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Using this command, global admins and Microsoft Teams service admins can access teams that they are not a member of.

When tabs are cloned, they are put into an unconfigured state. The first time you open them, you'll go through the configuration screen. If the person opening the tab does not have permission to configure apps, they will see a message explaining that the tab hasn't been configured.

## Examples

Creates a clone of a Microsoft Teams team with mandatory parameters

```sh
m365 teams team clone --teamId 15d7a78e-fd77-4599-97a5-dbb6372846c5 --displayName "Library Assist" --partsToClone "apps,tabs,settings,channels,members"
```

Creates a clone of a Microsoft Teams team with mandatory and optional parameters

```sh
m365 teams team clone --teamId 15d7a78e-fd77-4599-97a5-dbb6372846c5 --displayName "Library Assist" --partsToClone "apps,tabs,settings,channels,members" --description "Self help community for library" --classification "Library" --visibility "public"
```