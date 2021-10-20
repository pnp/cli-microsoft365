# v4 Upgrade Guidance

The v4 of CLI for Microsoft 365 introduces several breaking changes. To help you upgrade to the latest version of CLI for Microsoft 365, we've listed those changes along with any actions you may need to take.

## Changed default output to JSON

When building scripts using CLI for Microsoft 365, you typically use the JSON output to parse the response into an object and read its properties. In older version of CLI, the default output was text and you had to explicitly request JSON output mode using the `--output` option. As we see increased interest in building scripts with CLI for Microsoft 365, in v4 we switched the default output to JSON.

***What action do I need to take?***

You no longer need to use `--output json` in your scripts to use the JSON output mode. If you want to use the text output mode as default, you can configure it by executing `m365 cli config set --key output --value text`.

## In `spo listitem get` changed the `fields` option to `properties`

In the `spo listitem get` command we renamed the `fields` option to `properties` to better reflect the fact that the option allows you to specify not only the names of fields but also list item properties to retrieve.

***What action do I need to take?***

When using the `spo listitem get` command, use `--properties` instead of `--fields` to specify the list of properties to retrieve.

## Removed deprecated options that point to file paths

In v3.5 we introduced in the CLI for Microsoft 365 the ability to use the `@` token to pass the contents of a local file as the value of an option. With that, we no longer need separate options that let you specify a file path to pass file contents into a command. In this version, we removed the obsolete options in favor of options that support passing value both in-line and from a file using the `@` token.

***What action do I need to take?***

- in the `outlook mail send` command, replace the `--bodyContentsFilePath` option with `--bodyContents @file.ext`
- in the `spo theme set` command, replace the `--filePath` option with `--theme @file.ext`
- in the `teams team add` command, replace the `--templatePath` option with `--template @file.ext`

## Removed the `value` wrapper in some commands

In several commands, in the JSON output mode, we're inconsistently returning the data retrieved from Microsoft 365, wrapped in a `value` object. In this version, we removed the `value` wrapper from the output and aligned the commands with other commands in the CLI for Microsoft 365.

***What action do I need to take?***

When using the following commands in JSON output mode, update the logic processing the retrieved data to expect an array of objects without the `value` wrapper:

- `spo group list`
- `spo user list`
- `spo web list`
- `tenant service list`
- `tenant service message list`
- `tenant status list`

## Removed duplicate ID property in JSON output for `spo listitem` commands

In `spo listitem` commands, we were returning raw data from the SharePoint API, that contain the `Id` and `ID` properties. These properties lead to conflicts when trying to convert the output into an object in PowerShell. To avoid this issue, in v4 we removed the `ID` property from the output and kept the `Id` property.

***What action do I need to take?***

If you read list item IDs in scripts built using CLI for Microsoft 365, update your code to use the `Id` property.

## Separated `aad o365group user list` from `teams user list`

In the previous versions of CLI for Microsoft 365, the `teams user list` command was an alias of the `aad o365group user list` command and you could use either name in conjunction with Teams- or Office 365 groups-specific options. This led to confusion. As the CLI for Microsoft 365 evolved, we decided to split the commands and use only options specific to each workload.

***What action do I need to take?***

In the `aad o365group user list` command, you can only use the `--groupId` option which is now required. Use this command when working with Office 365 groups.

When working with Microsoft Teams teams, use the `teams user list` command for which you need to use the required `--teamId` option.

## Separated `aad o365group get` from `teams team get`

In the previous version of CLI for Microsoft 365, the `teams team get` command was an alias of the `aad o365group get` command. The `aad o365group get` command uses the `groups` endpoint from Microsoft Graph to retrieve information about Office 365 groups. With the release of the `teams` endpoint on Microsoft Graph, which returns specific information about Microsoft Teams teams, we decided to separate the two commands to allow you to retrieve the relevant information that you need to work with both workloads.

***What action do I need to take?***

If you use the `aad o365group get` command to retrieve information about Office 365 groups, you don't need to change anything. If you use the `teams team get` command to retrieve information about Teams, verify the new type of data returned by the command. Depending on your scripts, you might need to additionally run the `aad o365group get` command to get additional information about the underlying Office 365 group.
