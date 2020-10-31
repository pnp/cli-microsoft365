# planner task list

Lists Planner tasks for the currently logged in user

## Usage

```sh
m365 planner task list [options]
```

## Options

`-h, --help`
: output usage information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging.

`--debug`
: Runs command with debug logging.

## Examples

List tasks for the currently logged in user

```sh
m365 planner task list
```
