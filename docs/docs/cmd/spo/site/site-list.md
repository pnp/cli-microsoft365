# spo site list

Lists modern sites of the given type

## Usage

```sh
m365 spo site list [options]
```

## Options

`-h, --help`
: output usage information

`--type [type]`
: type of modern sites to list. Allowed values `TeamSite,CommunicationSite`, default `TeamSite`

`-f, --filter [filter]`
: filter to apply when retrieving sites

`--deleted`
: use this switch to only return deleted sites

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

`-h, --help`
: output usage information

`--type [type]`
: type of modern sites to list. Allowed values `TeamSite,CommunicationSite`, default `TeamSite`

`-f, --filter [filter]`
: filter to apply when retrieving sites

`--deleted`
: use this switch to only return deleted sites

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

## Remarks

Using the `-f, --filter` option you can specify which sites you want to retrieve. For example, to get sites with _project_ in their URL, use `Url -like 'project'` as the filter.

When using the text output type (default), the command lists only the values of the `Title`, and `Url` properties of the site. When setting the output type to JSON, all available properties are included in the command output.

## Examples

List all modern team sites in the currently connected tenant

```sh
m365 spo site list
```

List all modern team sites in the currently connected tenant

```sh
m365 spo site list --type TeamSite
```

List all modern communication sites in the currently connected tenant

```sh
m365 spo site list --type CommunicationSite
```

List all modern team sites that contain _project_ in the URL

```sh
m365 spo site list --type TeamSite --filter "Url -like 'project'"
```

List all deleted sites in the tenant you're logged in to

```sh
m365 spo site list --deleted
```
