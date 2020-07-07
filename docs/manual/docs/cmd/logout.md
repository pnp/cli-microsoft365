# logout

Log out from Office 365

## Usage

```sh
logout [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

The `logout` command logs out from Office 365 and removes any access and refresh tokens from memory

## Examples

Log out from Office 365

```sh
logout
```

Log out from Office 365 in debug mode including detailed debug information in the console output

```sh
logout --debug
```