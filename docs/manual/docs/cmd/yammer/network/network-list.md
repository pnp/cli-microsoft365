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
`--includeSuspended`|Include the networks in which the user is suspended.
`-o, --output [output]`|Output type. `json|text`. Default `text`

## Remarks

This command requires Yammer 'user_impersonation' permissions in Azure AD. 

The operations are executed in the context of the current logged in user. Certificate-based authentication with app_only permissions is not supported yet.  

## Examples

Returns the current user's networks.

```sh
yammer network list
```

Returns the current user's networks including the networks in which the user is suspsended.

```sh
yammer network list --includeSuspended
```