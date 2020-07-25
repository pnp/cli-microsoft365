# spo user list

Lists all the users within specific web

## Usage

```sh
m365 spo user list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the web to list the users from

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get list of users in web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo user list --webUrl https://contoso.sharepoint.com/sites/project-x
```
