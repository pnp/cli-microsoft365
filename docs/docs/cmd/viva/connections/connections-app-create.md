# viva connections app create

Creates a Viva Connections desktop app package to upload to Microsoft Teams

## Usage

```sh
m365 viva connections app create [options]
```

## Options

`--portalUrl <portalUrl>`
: The URL of the site to pin in Microsoft Teams. Must be a Communication site

`--appName <appName>`
: Name of the app to create, eg. `Contoso`. No longer than 30 characters

`--description <description>`
: Short description of the app. Displayed in the app's _About_ dialog. No longer than 80 characters

`--longDescription <longDescription>`
: Long description of the app. Displayed in the app's _About_ dialog. No longer than 4000 characters

`--privacyPolicyUrl [privacyPolicyUrl]`
: URL to your organization's privacy policy. Displayed in the app's _About_ dialog. Defaults to `https://privacy.microsoft.com/en-us/privacystatement` if not specified

`--termsOfUseUrl [termsOfUseUrl]`
: URL to your organization's terms of use. Displayed in the app's _About_ dialog. Defaults to `https://go.microsoft.com/fwlink/?linkid=2039674` if not specified

`--companyName <companyName>`
: Your organization's name. Displayed in the app's _About_ dialog

`--companyWebsiteUrl <companyWebsiteUrl>`
: Your organization's website URL. Displayed in the app's _About_ dialog

`--coloredIconPath <coloredIconPath>`
: Absolute or relative path to the color icon for your app

`--outlineIconPath <outlineIconPath>`
: Absolute or relative path to the outline icon for your app

`--accentColor [accentColor]`
: A HEX color to use in conjunction with and as a background for your outline icon. Defaults to `#40497E` if not specified

`--force`
: Specify, to overwrite the existing package file on disk

--8<-- "docs/cmd/_global.md"

## Remarks

If the specified portal URL doesn't exist, the command will a `404 - FILE NOT FOUND` error.

The specified portal URL must point to a valid Communication site. To get the list of Communication sites in your tenant, execute: `m365 spo site list --type CommunicationSite`.

The command generates a Microsoft Teams app package. App packages must meet specific requirements to be uploaded to Microsoft Teams. Specified attributes must not exceed their maximum length and the specified color and outline icons must be respectively 192x192px and 32x32px. For the latest list of requirements, see the links in the **More information** section at the end of this page. The generated app package will be written in the current working folder.

After creating the Viva Connections desktop app package, you need to upload it to your Microsoft Teams app catalog. You can do it either manually, or using the CLI by executing `m365 teams app publish --filePath ./contoso.zip`.

## Examples

Create a Viva Connections desktop app package

```sh
m365 viva connections app create --portalUrl https://contoso.sharepoint.com --appName Contoso --description "Contoso company app" --longDescription "Stay on top of what's happening at Contoso" --companyName Contoso --companyWebsiteUrl https://contoso.com --coloredIconPath icon-color.png --outlineIconPath icon-outline.png
```

## More information

- Add Viva Connections for Microsoft Teams desktop: [https://docs.microsoft.com/sharepoint/viva-connections](https://docs.microsoft.com/sharepoint/viva-connections?WT.mc_id=m365-15896-cxa)
- App manifest checklist: [https://docs.microsoft.com/microsoftteams/platform/concepts/deploy-and-publish/appsource/prepare/app-manifest-checklist](https://docs.microsoft.com/microsoftteams/platform/concepts/deploy-and-publish/appsource/prepare/app-manifest-checklist?WT.mc_id=m365-15896-cxa)
- Reference: Manifest schema for Microsoft Teams: [https://docs.microsoft.com/microsoftteams/platform/resources/schema/manifest-schema](https://docs.microsoft.com/microsoftteams/platform/resources/schema/manifest-schema?WT.mc_id=m365-15896-cxa)
