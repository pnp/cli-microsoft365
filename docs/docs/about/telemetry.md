# Telemetry

## General info

CLI for Microsoft 365 uses telemetry by default to track the usage within the project. Besides usage statistics, this gives us information to better understand how the CLI for Microsoft 365 is being used and where we can make improvements.

It is important to note that we **never** track personal information in our telemetry. For example, we track which command you are using and which command options, but the information passed to these options is never included in the telemetry.

To provide an overview, this is what we track when a command is executed:

- Name of the executed command
- Usage of command options (only usage, **not** data provided with the options)
- CLI for Microsoft 365 version
- Node.js version of your environment
- The type of shell you are using (PowerShell, Zsh, Cmd, Azure Cloud Shell, ...)
- Whether you are using a Docker container
- Whether you are using a CI/CD setup

## Example

An example is worth more than a thousand words. The following example illustrates which telemetry is being collected when executing a command.

Pretend we have CLI for Microsoft 365 version 6.0.0 installed within a Node.js v16.13.2 environment. Next, we open a PowerShell window and execute the following command.

```sh
m365 spo file add --webUrl "https://contoso.sharepoint.com/sites/project-x" --folder "/sites/project-x/Shared Documents" --path "C:\MS365.jpg" --contentType "Picture" --publish --publishComment "Lorem ipsum"
```

Executing this command will result in the following telemetry being collected:

Description | Telemetry data
----- | -----
Command name | spo file add
Command options[^1] | contentType: true<br />checkOut: false<br />checkInComment: false<br />approve: false<br />approveComment: false<br />publish: true<br />publishComment: true<br />query: false<br />output: json<br />verbose: false<br />debug: false
CLI for M365 version | 6.0.0
Node.js version | v16.13.2
Shell | pwsh.exe
Docker container | 
CI/CD setup | false

## Disable telemetry

!!! Note
    We offer the option to disable all telemetry within the project. However, we encourage you to leave it enabled as it helps us to understand the usage and impact of our work.

Run the following command to disable the telemetry.

```sh
m365 cli config set --key disableTelemetry --value true
```

## Re-enable telemetry

Run the following command to re-enable the telemetry.

```sh
m365 cli config reset --key disableTelemetry
```

[^1]: Note that we are only tracking the usage of optional options. Required options are always filled in and therefore there is no added value for us to include them in the telemetry.
