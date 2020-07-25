# teams app list

Lists apps from the Microsoft Teams app catalog or apps installed in the specified team

## Usage

```sh
m365 teams app list [options]
```

## Options

`-h, --help`
: output usage information

`-a, --all`
: Specify, to get apps from your organization only

`-i, --teamId [teamId]`
: The ID of the team for which to list installed apps

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

To list apps installed in the specified Microsoft Teams team, specify that team's ID using the `teamId` option. If the `teamId` option is not specified, the command will list apps available in the Teams app catalog.

## Examples

List all Microsoft Teams apps from your organization's app catalog only

```sh
m365 teams app list
```

List all apps from the Microsoft Teams app catalog and the Microsoft Teams store

```sh
m365 teams app list --all
```

List your organization's apps installed in the specified Microsoft Teams team

```sh
m365 teams app list --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55
```