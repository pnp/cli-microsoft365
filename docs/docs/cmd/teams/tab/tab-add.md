# teams tab add

Add a tab to the specified channel

## Usage

```sh
m365 teams tab add [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the team to where the channel exists

`-c, --channelId <channelId>`
: The ID of the channel to add a tab to

`--appId <appId>`
: The ID of the Teams app that contains the Tab

`--appName <appName>`
: The name of the Teams app that contains the Tab

`--contentUrl <contentUrl>`
: The URL used for rendering Tab contents

`--entityId [entityId]`
: A unique identifier for the Tab

`--removeUrl [removeUrl]`
: The URL displayed when a Tab is removed

`--websiteUrl [websiteUrl]`
: The URL for showing tab contents outside of Teams

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

The corresponding app must already be installed in the team.

## Examples
  
Add teams tab for website

```sh
m365 teams tab add --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --appId 06805b9e-77e3-4b93-ac81-525eb87513b8 --appName 'My Contoso Tab' --contentUrl 'https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView'
```

Add teams tab for website with additional configuration which is unknown

```sh
m365 teams tab add --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --appId 06805b9e-77e3-4b93-ac81-525eb87513b8 --appName 'My Contoso Tab' --contentUrl 'https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView' --test1 'value for test1'
```
