# Adding a command

Following article describes how to add a new command to the Office 365 CLI.

## Command files

Each command consists of three files:

- command implementation, located under **./src/o365/[service]/commands**, eg. *./src/o365/spo/commands/connect.ts*
- command unit tests, located under **./src/o365/[service]/commands**, eg. *./src/o365/spo/commands/connect.spec.ts*
- command documentation page, located under **./docs/manual/docs/cmd/[service]**, eg. *./docs/manual/docs/cmd/spo/connect.md*

Additionally, each command is listed in:

- list of all commands for the given service, located in **./src/o365/[service]/commands.ts**, eg. *./src/o365/spo/commands.ts*
- the documentation table of contents, located in **./docs/manual/mkdocs.yml**

## Add new files

Commands are organized by the Office 365 service, such as SharePoint Online (spo), that they apply to. Before building your command, find the right folder corresponding with your command in the project structure.

### Create new files

In the **./src/o365/[service]/commands** folder, create two files for your command: **my-command.ts** for the command implementation, and **my-command.spec.ts** for the unit tests.

### Define command name constant

In the **./src/o365/[service]/commands.ts** file, define a constant with your command's name including the service prefix. You will use this constant to refer to the command in its implementation, unit tests, help, etc.

### Add the command manual page

In the **./docs/manual/docs/cmd/[service]** folder, create new file for your command's help page: **my-command.md**. Next, open the **./docs/manual/mkdocs.yml** file and add the reference to the **my-command.md** file in the table of contents.

> The table of contents is organized alphabetically so that users can quickly find the command they are looking for.

## Implement command

Each command in the Office 365 CLI is defined as a class extending the [Command](../../src/Command.ts) base class. At minimum a command must define `name`, `description`, `action` and `help`:

```ts
import config from '../../../config';
import commands from '../commands';
import Command, {
  CommandAction,
  CommandHelp
} from '../../../Command';
import appInsights from '../../../appInsights';

const vorpal: Vorpal = require('../../../vorpal-init');

class SpoMyCommand extends Command {
  public get name(): string {
    return commands.MYCOMMAND;
  }

  public get description(): string {
    return 'My command';
  }

  public get action(): CommandAction {
    return function (this: CommandInstance, args: {}, cb: () => void) {
      appInsights.trackEvent({
        name: commands.MYCOMMAND
      });

      // command implementation goes here

      cb(); // notify that the command completed
    };
  }

  public help(): CommandHelp {
    return function (args: any, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.MYCOMMAND).helpInformation());
      log(
        `  Remarks:

    Here are some additional considerations when using this command.

  Examples:

    ${chalk.grey(config.delimiter)} ${commands.MYCOMMAND}
      example one of using the command
`);
    };
  }
}

module.exports = new SpoMyCommand();
```

### Tracking command usage

The Office 365 CLI tracks usage of the different commands using Azure Application Insights. Each command action should begin with logging its usage, by including:

```ts
class SpoMyCommand extends Command {
  // ...

  public get action(): CommandAction {
    return function (this: CommandInstance, args: {}, cb: () => void) {
      appInsights.trackEvent({
        name: commands.MYCOMMAND
      });

      // ...
    };
  }

  // ...
}
```

If your command has additional parameters, you should include them in the telemetry as well:

```ts
class SpoMyCommand extends Command {
  // ...

  public get action(): CommandAction {
    return function (this: CommandInstance, args: {}, cb: () => void) {
      appInsights.trackEvent({
        name: commands.TENANT_CDN_GET,
        properties: {
          cdnType: cdnTypeString,
          verbose: verbose.toString()
        }
      });

      // ...
    };
  }

  // ...
}
```

> **Important:** if your command requires URLs or other user-defined strings, you **should not** include them in the telemetry as these strings might include personal or confidential information that we shouldn't have access to.

### Notifying when command action executed

When executing the command completed, you should notify the CLI of it, by calling the callback method which is the last argument in the function returned in the **action** method:

```ts
class SpoMyCommand extends Command {
  // ...

  public get action(): CommandAction {
    return function (this: CommandInstance, args: {}, cb: () => void) {
      appInsights.trackEvent({
        name: commands.MYCOMMAND
      });

      // command implementation goes here

      cb(); // notify that the command completed
    };
  }

  // ...
}
```

> **Important:** if you don't call the callback method, the CLI won't exit to the command prompt and users won't be able to run additional commands.

### Defining command help

Vorpal, the engine upon which the Office 365 CLI is built, renders rudimentary help for each command. In the Office 365 CLI we extend this basic information with additional remarks and examples to help users work with the CLI.

When building command help, you can get the standard help from Vorpal by calling: `vorpal.find('commandname').helpInformation()`. Using the `log` method you can include additional information.

```ts
class SpoMyCommand extends Command {
  // ...
  public help(): CommandHelp {
    return function (args: any, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.MYCOMMAND).helpInformation());
      log(
        `  Remarks:

    Here are some additional considerations when using this command.

  Examples:

    ${chalk.grey(config.delimiter)} ${commands.MYCOMMAND}
      example one of using the command
`);
    };
  }
}
```

> To emphasize important information or references to other commands, you can use **chalk**. See the implementations of existing commands to see how it's used.

### Export command class instance

Finish the implementation of your command, by exporting the instance of the command class:

```ts
module.exports = new SpoMyCommand();
```

On runtime, Office 365 CLI iterates through all JavaScript files in the **o365** folder and registers all exported command classes as commands in the CLI.

### Additional command capabilities

When building Office 365 CLI commands, you can use additional features such as optional and required arguments, autocomplete or validation. For more information see the [Vorpal command API documentation](https://github.com/dthree/vorpal/wiki/API-%7C-vorpal.command).

## Implement unit tests

Each command must by accompanied by a set of unit tests. We aim for 100% code and branch coverage in every command file.

> To see the current code coverage, run `npm test`. Once testing completes, open the **./coverage/lcov-report/index.html** file in the web browser and browser through the different project files.

See the existing test files to get a better understanding of how they are structured and how different elements such as objects or web requests are mocked.

Once you're done with your unit tests, run `npm test` to verify that you're covering all code and branches of your command with your unit tests.

## Write help page

Each command has a corresponding manual page. The contents of this page are almost identical to the help implemented in the command itself. This way, users working with the CLI can get the help directly inside the CLI, while users interested in the capabilities of the CLI, can browse through the help pages published on the Internet.

Start filling the help page contents by starting the Office 365 CLI and requesting help for your command:

```sh
o365$ help spo my-command
```

Copy the output of the command and use as a starting point for creating the documentation page. The main difference between the help displayed in the CLI and the manual page is the formatting. In the command line, the CLI uses chalk to emphasize information. The manual uses Markdown to format the output. To maintain consistency, refer to other manual pages to see how they are structured and how the information is presented.

## That's it

If the project is still building, your command is working as expected, all unit tests are passing, you have 100% code coverage on your command file and the documentation is in place, you're ready to [submit a PR](./submitting-pr.md).