# Building the command

To start writing the logic for the command, you will need to create a new TypeScript file. All the command services can be found in the folder `src/m365`. Here you will find all the services available within the CLI for Microsoft 365. Each service contains a subfolder named `commands` where all the service commands are located. For our example issue, mentioned [here](./step-by-step-guide.md#new-command-get-site-group), that will be `src/m365/spo/commands`.

!!! tip

    When the service has a lot of commands within the `commands` folder then they will be split up into subfolders. E.g. `src/m365/spo/commands/group`

When you are in the correct folder you can create two new files. Your command file `group-get.ts` and a file for the unit tests `group-get.spec.ts`.

## Minimum command file 

With our two new files created, we can start working on our `group-get.ts` file. Each command in the CLI for Microsoft 365 is defined as a class extending the [Command](https://github.com/pnp/cli-microsoft365/blob/main/src/Command.ts) base class. At minimum a command must define `name`, `description`, and `commandAction`:

```ts
import commands from '../../commands';
import { Logger } from '../../../../cli/Logger';
import SpoCommand from '../../../base/SpoCommand';

class SpoGroupGetCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_GET;
  }

  public get description(): string {
    return 'Gets site group';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information for group in site ...`);
    }

    // Command implementation goes here
  }
}

module.exports = new SpoGroupGetCommand();
```

Depending on your command and the service for which you're building the command, there might be a base class that you can use to simplify the implementation. For example for SPO, you can inherit from the [SpoCommand](https://github.com/pnp/cli-microsoft365/blob/main/src/m365/base/SpoCommand.ts) base class. This class contains several helper methods to simplify your implementation.

### Include command name

When you create the minimum file, you'll get an error about a none existing type within `commands`. This is correct because we haven't defined the name of the command yet. Let's add this to the `commands` export located in `src/m365/spo/commands.ts`.

```ts title="src/m365/spo/commands.ts"
const prefix: string = 'spo';

export default {
  // ...
  GROUP_GET: `${prefix} group get`,
  // ...
}
```

Next up, to enhance our command with options, validators, telemetry, ... There are a bunch of methods already available for you.

## Defining the options

When the command requires options to be passed along, we will define them in the interface `Options`. This interface extends from our GlobalOptions where the common options `query`, `output`, `debug`, and `verbose` are defined. When an option is optional let's make sure that it's also optional in the interface.

We will also define the options in the method `#initOptions`. Here we pass along the option name, as a possible abbreviation for the option to `this.options` object. In some occasions, the option will always require a pre-defined input. When this is the case, we can define them under the property `autocomplete`.

!!! tip

    Required options are denoted as `--required <required>` and optional options are denoted as `--optional [optional]` 

```ts title="group-get.ts"
// ...
import GlobalOptions from '../../../../GlobalOptions';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: number;
  name?: string;
  associatedGroup?: string;
}

class SpoGroupGetCommand extends SpoCommand {
  // ...

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--name [name]'
      },
      {
        option: '--associatedGroup [associatedGroup]',
        autocomplete: ['Owner', 'Member', 'Visitor']
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information for group in site at ${args.options.webUrl}...`);
    }
    
    // Command implementation goes here
  }
}
```

## Option validation

The options that are passed along won't always be correct from the first, so instead of passing faulty values to the required API, we can write option validation that runs before `commandAction` is executed. This can be done in the method `initValidators`. Conditions can be written here to validate the option values and return an error when it's faulty. Once again, there are already several validation methods you can make use of to check some common options. e.g. `validation.isValidSharePointUrl(...)`.

```ts title="group-get.ts"
// ...
import { validation } from '../../../../utils/validation';

class SpoExampleListCommand extends Command {
  constructor() {
    super();
  
    // ...
    this.#initValidators();
  }

  // ...

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && isNaN(args.options.id)) {
          return `Specified id ${args.options.id} is not a number`;
        }

        if (args.options.associatedGroup && ['owner', 'member', 'visitor'].indexOf(args.options.associatedGroup.toLowerCase()) === -1) {
          return `${args.options.associatedGroup} is not a valid associatedGroup value. Allowed values are Owner|Member|Visitor.`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }
}
```

## Option sets

Option sets are used to ensure that one option contains a value from a set of options. When no option is used, the command will return an error and the same goes when multiple of these options are used. To make use of the option sets, you can use the method `#initOptionSets`.

```ts title="group-get.ts"
class SpoExampleListCommand extends Command {
  constructor() {
    super();
  
    // ...
    this.#initOptionSets();
  }

  // ...

  #initOptionSets(): void {
    this.optionSets.push(['id', 'name', 'associatedGroup']);
  }
}
```

## Telemetry

The CLI for Microsoft 365 tracks the usage of the different commands using Azure Application Insights. By default, for each command the CLI tracks its name and whether it's been executed in debug/verbose mode or not. If your command has additional properties that should be included in the telemetry, you can define them by implementing the `#initTelemetry` method and adding your properties to `this.telemetryProperties` object.

```ts title="group-get.ts"
class SpoExampleListCommand extends Command {
  constructor() {
    super();
  
    // ...
    this.#initTelemetry();
  }

  // ...

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        email: typeof args.options.email !== 'undefined',
        type: typeof args.options.type !== 'undefined'
      });
    });
  }
}
```

## Command action

After everything is written for our options, we can start to write the logic required to execute the command. This will be done, as mentioned before, in the method `commandAction`. The command will start with a verbose message explaining what we are about to do and then we start writing the command logic. Here you can write several new methods to be called in `commandAction` to keep the code a bit tidier.

When writing your code, there are a few pointers to keep in mind:

- It's recommended to add some verbose logging along the path of your command. This can keep the user informed about what you are doing.
- Async tasks will be written with async/await.
- When an endpoint errors, make sure that the output is returned to the user.

```ts title="group-get.ts"
class SpoExampleListCommand extends Command {

  // ...

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information for group in site at ${args.options.webUrl}...`);
    }

    // ...
    // Command logic
    // ...

    try {
      const groupInstance = await request.get(requestOptions);
      logger.log(groupInstance);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}
```

After this, the new command will be fully functional. During the development, it can be useful to have `npm run watch` running in the background. This builds the entire project first. After this, a watcher will make sure that every time a file is saved, an incremental build is triggered. This means that not the entire project is rebuilt but only the changed files. That way you can easily apply new changes to the command and test it out locally.

> In the end, your command file will look something like this: [group-get.ts](https://github.com/pnp/cli-microsoft365/blob/main/src/m365/spo/commands/group/group-get.ts)

## Running it locally

Before creating a PR you should test your code locally. This will help you to catch bugs, errors, and performance issues early on and ensures that the code is functioning as intended before it is made available to real users. You can execute `npm run watch` to start a live watcher. This will build the entire project first and after this, a watcher will make sure that every time a file is saved, an incremental build is triggered. This means that not the entire project is rebuilt but only the changed files. This makes it easy to do some quick changes and test them immediately after you have saved them. 

If this command fails, be sure to check if your environment has been set up correctly following the guidelines of ["Setting up your local project"](../environment-setup.md#setting-up-your-local-project)
