# spo site set

Updates properties of the specified site

## Usage

```sh
m365 spo site set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: The URL of the site collection to update

`--classification [classification]`
: The new classification for the site collection

`--disableFlows [disableFlows]`
: Set to `true` to disable using Microsoft Flow in this site collection

`--isPublic [isPublic]`
: Set to `true` to make the group linked to the site public or to `false` to make it private

`--owners [owners]`
: Comma-separated list of users to add as site collection administrators

`--shareByEmailEnabled [shareByEmailEnabled]`
: Set to true to allow to share files with guests and to false to disallow it

`--siteDesignId [siteDesignId]`
: Id of the custom site design to apply to the site

`--title [title]`
: The new title for the site collection

`--sharingCapability [sharingCapability]`
: The sharing capability for the site. Allowed values:  `Disabled`, `ExternalUserSharingOnly`, `ExternalUserAndGuestSharing`, `ExistingExternalUserSharingOnly`.

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

If the specified url doesn't refer to an existing site collection, you will get a `404 - "404 FILE NOT FOUND"` error.

The `isPublic` property can be set only on groupified site collections. If you try to set it on a site collection without a group, you will get an error.

When setting owners, the specified owners will be added to the already configured owners. Existing owners will not be removed.

## Examples

Update site collection's classification

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --classification MBI
```

Reset site collection's classification.

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --classification
```

Disable using Microsoft Flow on the site collection

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --disableFlows true
```

Update the visibility of the Microsoft 365 group behind the specified groupified site collection to public

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --isPublic true
```

Update site collection's owners

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --owners "john@contoso.onmicrosoft.com,steve@contoso.onmicrosoft.com"
```

Allow sharing files in the site collection with guests

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --shareByEmailEnabled true
```

Apply the specified site ID to the site collection

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --siteDesignId "eb2f31da-9461-4fbf-9ea1-9959b134b89e"
```

Update site collection's title

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --title "My new site"
```

Restrict external sharing to already available external users only

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --sharingCapability ExistingExternalUserSharingOnly
```
