# spo knowledgehub remove

Removes the Knowledge Hub Site setting for your tenant

## Usage

```sh
m365 spo knowledgehub remove [options]
```

## Options

`-h, --help`
: output usage information

`--confirm`
: Do not prompt for confirmation before removing the Knowledge Hub Site setting for your tenant

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Removes the Knowledge Hub Site setting for your tenant

```sh
m365 spo knowledgehub remove
```

Removes the Knowledge Hub Site setting for your tenant without confirmation

```sh
m365 spo knowledgehub remove --confirm
```
