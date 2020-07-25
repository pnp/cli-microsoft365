# status

Shows Microsoft 365 login status

## Usage

```sh
m365 status [options]
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

If you are logged in to Microsoft 365, the `status` command will show you information about the user or application name used to sign in and the details about the stored refresh and access tokens and their expiration date and time when run in debug mode.

## Examples

Show the information about the current login to the Microsoft 365

```sh
m365 status
```
