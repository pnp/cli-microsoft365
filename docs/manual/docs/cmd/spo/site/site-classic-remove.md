# spo site classic remove

Remove classic site

## Usage

```sh
spo site classic remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url`| url of the site to remove
`--skipRecycleBin`|set to directly remove the site without moving it to the Recycle Bin
`--fromRecycleBin`|set to remove the site from the Recycle Bin
`--wait`|Wait for the site to be removed before completing the command
`--confirm`|Don't prompt for confirming removing the file
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks
    To remove a classic site, you have to first connect to a tenant admin
    site using the [spo connect](../connect.md) command.
        
    Deleting and creating classic site collections is by default asynchronous and depending on the current state of Office 365, might take up to few minutes. If you're building a script with steps that require the site to be fully deleted, you should use the '--wait' flag. When using this flag, the command will keep running until it received confirmation from Office 365 that the site
    has been fully deleted.

## Examples

Remove the site based on URL, and place it in the recycle bin

```sh
spo site classic remove --url https://contoso.sharepoint.com/sites/demosite
```

Remove the site based on URL permanently 

```sh
spo site classic remove --url https://contoso.sharepoint.com/sites/demosite --skipRecycleBin 
```

Remove the site based on URL from the recycle bin

```sh
spo site classic remove --url https://contoso.sharepoint.com/sites/demosite --fromRecycleBin 
```

Remove the site based on URL permanently and wait for completion

```sh
spo site classic remove --url https://contoso.sharepoint.com/sites/project-x --wait --skipRecycleBin
```