# teams team add

Adds a new Microsoft Teams team

## Usage

```sh
m365 teams team add [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name [name]`
: Display name for the Microsoft Teams team. Required if `templatePath` not supplied

`-d, --description [description]`
: Description for the Microsoft Teams team. Required if `templatePath` not supplied

`--templatePath [templatePath]`
: Local path to the file containing the template. If `name` or `description` are supplied, these take precedence over the template values

`--wait`
: Wait for the team to be provisioned before completing the command

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

If you want to add a Team to an existing Microsoft 365 Group use the [aad o365group teamify](../../aad/o365group/o365group-teamify.md) command instead.

This command will return different responses based on the presence of the `--wait` option. If present, the command will return a `group` resource in the response. If not present, the command will return a `teamsAsyncOperation` resource in the response.

## Examples

Add a new Microsoft Teams team

```sh
m365 teams team add --name 'Architecture' --description 'Architecture Discussion'
```

Add a new Microsoft Teams team using a template

```sh
m365 teams team add --name 'Architecture' --description 'Architecture Discussion' --templatePath 'template.json'
```

Add a new Microsoft Teams team using a template and wait for the team to be provisioned

```sh
m365 teams team add --name 'Architecture' --description 'Architecture Discussion' --templatePath 'template.json' --wait
```

## More information

- Get started with Teams templates: [https://docs.microsoft.com/en-us/MicrosoftTeams/get-started-with-teams-templates](https://docs.microsoft.com/en-us/MicrosoftTeams/get-started-with-teams-templates)
- group resource type: [https://docs.microsoft.com/en-gb/graph/api/resources/group?view=graph-rest-beta](https://docs.microsoft.com/en-gb/graph/api/resources/group?view=graph-rest-beta)
- teamsAsyncOperation resource type: [https://docs.microsoft.com/en-gb/graph/api/resources/teamsasyncoperation?view=graph-rest-beta](https://docs.microsoft.com/en-gb/graph/api/resources/teamsasyncoperation?view=graph-rest-beta)
