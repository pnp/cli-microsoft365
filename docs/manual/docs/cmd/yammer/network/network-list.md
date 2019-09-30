# Yammer network list

Returns a list of networks to which the current user has access

## Usage

```sh
yammer network list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--includeSuspended`|Include the networks the user is suspended.
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

This command requires Yammer 'user_impersonation' permissions in Azure AD. 

The operations are executed in the context of the current logged in user. Certificate-based authentication with app_only permissions is not supported yet.  

## Examples

Returns a list of networks to which the current user has access.

```sh
yammer network list
```

Returns a list of networks to which the current user has access including the networks the user is suspended.

```sh
yammer network list --includeSuspended
```