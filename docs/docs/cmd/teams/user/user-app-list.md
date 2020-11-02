# teams user app list

List the apps installed in the personal scope of the specified user

## Usage

```sh
m365 teams user app list [options]
```

## Options

`-h, --help`
: output usage information

`--userId [userId]`
: The ID of the user to get the apps from. Specify `userId` or `userName` but not both.

`--userName [userName]`
: The UPN of the user to get the apps from. Specify `userId` or `userName` but not both.

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

List the apps installed in the personal scope of the specified user using its ID

```sh
m365 teams user app list --userId 4440558e-8c73-4597-abc7-3644a64c4bce
```

List the apps installed in the personal scope of the specified user using its UPN

```sh
m365 teams user app list --userName admin@contoso.com
```
