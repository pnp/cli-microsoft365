# spo site rename

Renames the URL and title of a site collection

## Usage

```sh
m365 spo site rename [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: The URL of the site to rename

`--newSiteUrl <newSiteUrl>`
: New URL for the site collection

`--newSiteTitle [newSiteTitle]`
: New title for the site

`--suppressMarketplaceAppCheck`
: Suppress marketplace app check

`--suppressWorkflow2013Check`
: Suppress 2013 workflow check

`--wait`
: Wait for the job to complete

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you must have permissions to access the tenant admin site.

## Remarks

Renaming site collections is by default asynchronous and depending on the current state of Microsoft 365, might take up to few minutes. If you're building a script with steps that require the operation to complete fully, you should use the `--wait` flag. When using this flag, the `spo site rename` command  will keep running until it receives confirmation from Microsoft 365 that the site rename operation has completed.

## Examples

Starts the rename of the site collection with name "samplesite" to "renamed" without modifying the title

```sh
m365 spo site rename --siteUrl http://contoso.sharepoint.com/samplesite --newSiteUrl http://contoso.sharepoint.com/renamed
```

Starts the rename of the site collection with name "samplesite" to "renamed" modifying the title of the site to "New Title"

```sh
m365 spo site rename --siteUrl http://contoso.sharepoint.com/samplesite --newSiteUrl http://contoso.sharepoint.com/renamed --newSiteTitle "New Title"
```

Renames the specified site collection and waits for the operation to complete

```sh
m365 spo site rename --siteUrl http://contoso.sharepoint.com/samplesite --newSiteUrl http://contoso.sharepoint.com/renamed --newSiteTitle "New Title" --wait
```
