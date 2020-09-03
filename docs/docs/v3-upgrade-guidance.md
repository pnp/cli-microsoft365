# v3 Upgrade Guidance

We recently announced that the next version of the Office 365 CLI will be v3 and is planned to be released on September 6, 2020.

With this release there are a number of breaking changes, so to help you make the transition from `Office 365 CLI v2` to `CLI for Microsoft 365 v3`, we have listed those changes, why they have been made and what action you may need to take.

## New package name

With the move to the new name, a new package has been published to npm with the name, `@pnp/cli-microsoft365`

As with v2, we will still provide a few variations of package.

Use `@pnp/cli-microsoft365@next` for the latest beta version, we release beta versions on a regular basis, for you to try new features and fixes before the monthly stable release.

Use `@pnp/cli-microsoft365` or `@pnp/cli-microsoft365@latest` for the latest stable version, which we release every month.

### Why `CLI for Microsoft 365` and not `Microsoft 365 CLI`?

This is a community-owned product, as Microsoft uses product names that begin with _Microsoft_ and _Microsoft 365_ for products they own, we did not want to cause confusion between the two.

***What action do I need to take?***

If you have scripts using `@pnp/office365-cli` then we would encourage you to update your installed package to use the new package name, `@pnp/cli-microsoft365`, this will ensure that you get all the latest updates and patches.

!!! attention
    The `@pnp/office365-cli` package will still remain on npm, but will be deprecated following the release of v3. It will not receive any updates or fixes, therefore we would strongly recommend that you should plan to upgrade to the new package.

## Removal of the o365 command alias

In Office 365 CLI v2, we provided two command aliases, `o365` and later, `m365`. With the release of v3, we will be removing the `o365` command alias.

***What action do I need to take?***

After upgrading to v3, `@pnp/cli-microsoft365`, any scripts that use the `o365` alias will need to be updated to use `m365` alias instead.

## Renaming of CLI environment variables

In Office 365 CLI v2, we provided two environment variables, `OFFICE365CLI_AADAPPID` and `OFFICE365CLI_TENANT`. These environment variable could be updated to point the CLI to your own custom Azure Active Directory identity to use for logging into your Microsoft 365 tenant instead of using the multi-tenant PnP Management Shell identity used as the default identity.

Find out more about this feature: [Using your own Azure AD identity](user-guide/using-own-identity.md)

These environment variables have been renamed following the rename of the package.

***What action do I need to take?***

After upgrading to v3, `@pnp/cli-microsoft365`, any scripts that make use of a custom Azure AD identity for authentication will need to update the setting of the environment variables to use the new names, `CLIMICROSOFT365_AADAPPID` and `CLIMICROSOFT365_TENANT`.

## Removal of the global `--pretty` option

In Office 365 CLI v2, we introduced a global option to prettify the JSON output, this helped to make the output more readable when working with large JSON responses. Rather than having this feature as an explicit option, we have decided to remove this option and make the default JSON output prettified.

***What action do I need to take?***

After upgrading to v3, `@pnp/cli-microsoft365`, any scripts that use the `--pretty` option will need to be updated and you should remove the option from the command to ensure the it works as expected.

## Removal of immersive mode

In Office 365 CLI v2, we provided an immersive mode, which you could enter by executing the `o365` or `m365` command with no options passed, however we made the decision to remove the immersive mode from the CLI.

For a couple of reasons, the first being that the usage of this mode was very low, secondly, removing this mode simplifies our code base and removes a dependency on the Vorpal library which the CLI uses behind the scenes, making the CLI much more maintainable going forwards.

***What action do I need to take?***

There is no action for you to take.

## Removal of `--outputFile` option

In Office 365 CLI v2, we provided an option called `--outputFile` on certain commands. This option provided a way of saving the output of the command to a local file, this was particularly useful when using the immersive mode where it was not possible to use any shell commands. As we have removed the immersive mode in v3, the requirement for a specific option is no longer required and it is more convenient to use shell commands for this purpose.

***What action do I need to take?***

After upgrading to v3, `@pnp/cli-microsoft365`, any scripts that use the `--outputFile` option should be updated. Remove the option and replace with a shell command to send the output to a file, for example, `m365 aad o365group report activitycounts --period D7 --output json --outputFile "o365groupactivitycounts.json"` becomes `m365 aad o365group report activitycounts --period D7 --output json > "o365groupactivitycounts.json"`

## Removal of deprecated aliases

In Office 365 CLI v2, we had a number of aliases that had been deprecated but remained within the CLI for backwards compatibility. In v3, we have taken the decision to remove these aliases, they are listed below with their replacement aliases.

- `consent` > `m365 cli consent`
- `--reconsent` > `m365 cli reconsent`
- `--completion:clink:generate` > `m365 cli completion clink update`
- `--completion:sh:generate` > `m365 cli completion sh update`
- `--completion:sh:setup` > `m365 cli completion sh setup`
- `accesstoken get` > `m365 util accesstoken get`

***What action do I need to take?***

After upgrading to v3, `@pnp/cli-microsoft365`, any scripts that use any of the deprecated aliases option should be updated to use their replacement aliases.

## Deprecation of support for options with spaces without quotes

In Office 365 CLI v2, it was possible to specify options with spaces without quotes. For example, `m365 spo site add --title My new site`, going forwards into v3 we will not support this.

***What action do I need to take?***

After upgrading to v3, `@pnp/cli-microsoft365`, any scripts that use option values containing spaces, must be  updated to wrap the option value in quotes, for example, `m365 spo site add --title My new site` will become `m365 spo site add --title "My new site"`, this will ensure that the command works as expected.

## Renaming of `--query` option on `spo search` command

In Office 365 CLI v2, we identified that on the `spo search` command the option used to pass in the search query was called `--query`, unfortunately when we introduced the ability to use JMESPath queries we introduced a global option with the same name and this caused the command to break.

***What action do I need to take?***

After upgrading to v3, `@pnp/cli-microsoft365`, rename the `--query` option to `--queryText` when using the `spo search` command.

## Removal of `-h` short option on `spo contenttype field set` and `spo hidedefaultthemes set` commands

In Office 365 CLI v2, to return usage examples of a command in your shell, you can use the `--help` option. We wanted to introduce a short version of this option however we had identified that two commands already had the `-h` short option implemented.

We have taken the step to make `-h` short option a reserved global option in v3, which makes the CLI consistent with other CLIs such as the Azure CLI.

***What action do I need to take?***

After upgrading to v3, `@pnp/cli-microsoft365`, any scripts that use the the short options on spo contenttype field set and spo hidedefaultthemes set commands should be updated to use their longer named options, `--hidden` and `--hideDefaultTheme`, respectively.