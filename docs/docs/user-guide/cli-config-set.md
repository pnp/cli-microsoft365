# Configuring the CLI for Microsoft 365

When running your commands or your scripts, you can specify settings regarding the behavior of the CLI.

## Setting example

```sh
m365 cli config set --key showHelpOnFailure --value true
```

## Available settings

| Setting name                | Definition                                     | Default value  |
|-----------------------------|------------------------------------------------|----------------|
| showHelpOnFailure           | Displays command help when the execution fails | false          |
