# spo site remove

Removes the specified site

## Usage

```sh
m365 spo site remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: URL of the site to remove

`--skipRecycleBin`
: Set to directly remove the site without moving it to the Recycle Bin

`--fromRecycleBin`
: Set to remove the site from the Recycle Bin

`--wait`
: Wait for the site to be removed before completing the command

`--confirm`
: Don't prompt for confirming removing the site

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

Deleting a site collection is by default asynchronous and depending on the current state of Microsoft 365, might take up to few minutes. If you're building a script with steps that require the site to be fully deleted, you should use the `--wait` flag. When using this flag, the `spo site remove` command will keep running until it received confirmation from Microsoft 365 that the site has been fully deleted.

## Examples

Remove the specified site and place it in the Recycle Bin

```sh
m365 spo site remove --url https://contoso.sharepoint.com/sites/demosite
```

Remove the site without moving it to the Recycle Bin

```sh
m365 spo site remove --url https://contoso.sharepoint.com/sites/demosite --skipRecycleBin
```

Remove the previously deleted site from the Recycle Bin

```sh
m365 spo site remove --url https://contoso.sharepoint.com/sites/demosite --fromRecycleBin
```

Remove the site without moving it to the Recycle Bin and wait for completion

```sh
m365 spo site remove --url https://contoso.sharepoint.com/sites/demosite --wait --skipRecycleBin
```
