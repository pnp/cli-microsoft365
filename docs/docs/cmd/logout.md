# logout

Log out from Microsoft 365

## Usage

```sh
m365 logout [options]
```

## Options

`-h, --help`
: output usage information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

The `logout` command logs out from Microsoft 365 and removes any access and refresh tokens from memory

## Examples

Log out from Microsoft 365

```sh
m365 logout
```

Log out from Microsoft 365 in debug mode including detailed debug information in the console output

```sh
m365 logout --debug
```