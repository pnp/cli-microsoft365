# spo externaluser list

Lists external users in the tenant

## Usage

```sh
m365 spo externaluser list [options]
```

## Options

`-h, --help`
: output usage information

`-f, --filter [filter]`
: Limits the results to only those users whose first name, last name or email address begins with the text in the string, using a case-insensitive comparison

`-p, --pageSize [pageSize]`
: Specifies the maximum number of users to be returned in the collection. The value must be less than or equal to `50`

`-i, --position [position]`
: Use to specify the zero-based index of the position in the sorted collection of the first result to be returned

`-s, --sortOrder [sortOrder]`
: Specifies the sort results in Ascending or Descending order on the `SPOUser.Email` property should occur. Allowed values `asc|desc`. Default `asc`

`-u, --siteUrl [siteUrl]`
: Specifies the site to retrieve external users for. If no site is specified, the external users for all sites are returned

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

List all external users from the current tenant. Show the first batch of 50 users.

```sh
m365 spo externaluser list --pageSize 50 --position 0
```

List all external users from the current tenant whose first name, last name or email address
begins with `Vesa`. Show the first batch of 50 users.

```sh
m365 spo externaluser list --pageSize 50 --position 0 --filter Vesa
```

List all external users from the specified site. Show the first batch of 50 users.

```sh
m365 spo externaluser list --pageSize 50 --position 0 --siteUrl https://contoso.sharepoint.com
```

List all external users from the current tenant. Show the first batch of 50 users sorted descending
by e-mail.

```sh
m365 spo externaluser list --pageSize 50 --position 0 --sortOrder desc
```