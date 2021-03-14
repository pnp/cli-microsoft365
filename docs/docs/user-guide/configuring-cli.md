# Configuring CLI for Microsoft 365

You can configure CLI for Microsoft 365 to your personal preferences using its settings. Settings are stored on the disk in the current user's folder: `C:\Users\user\.config\configstore\cli-m365-config.json` on Windows and `/Users/user/.config/configstore/cli-m365-config.json` on macOS. The configuration file is created when you set the settings for the first time.

To reset settings to their default values, remove them from the configuration file or remove the whole configuration file.

## Configuring settings

You can configure the specific setting using the `cli config set` command. For example, to configure CLI to automatically show help when executing a command failed, execute:

```sh
m365 cli config set --key showHelpOnFailure --value true
```

## Available settings

Following is the list of configuration settings available in CLI for Microsoft 365.

Setting name|Definition|Default value
------------|----------|-------------
`showHelpOnFailure`|Automatically display help when executing a command failed|`true`
`output`|Defines the default output when issuing a command|`text`
