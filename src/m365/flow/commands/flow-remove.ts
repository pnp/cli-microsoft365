import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../Command';
import request from '../../../request';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import Utils from '../../../Utils';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  name: string;
  asAdmin?: boolean;
  confirm?: boolean;
}

class FlowRemoveCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Flow';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.asAdmin = typeof args.options.asAdmin !== 'undefined';
    telemetryProps.confirm = typeof args.options.confirm !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Removing Microsoft Flow ${args.options.name}...`);
    }

    const removeFlow: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
        resolveWithFullResponse: true,
        headers: {
          accept: 'application/json'
        },
        json: true
      };

      request
        .delete(requestOptions)
        .then((rawRes: any): void => {
          // handle 204 and throw error message to cmd when invalid flow id is passed
          // https://github.com/pnp/cli-microsoft365/issues/1063#issuecomment-537218957
          if (rawRes.statusCode === 204) {
            cmd.log(vorpal.chalk.red(`Error: Resource '${args.options.name}' does not exist in environment '${args.options.environment}'`));
            cb();
          }
          else {
            if (this.verbose) {
              cmd.log(vorpal.chalk.green('DONE'));
            }
            cb();
          }
        }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
    };
    if (args.options.confirm) {
      removeFlow();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Microsoft Flow ${args.options.name}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeFlow();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'The name of the Microsoft Flow to remove'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment to which the Flow belongs'
      },
      {
        option: '--asAdmin',
        description: 'Set, to remove the Flow as admin'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirmation'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required option name missing';
      }

      if (!Utils.isValidGuid(args.options.name)) {
        return `${args.options.name} is not a valid GUID`;
      }

      if (!args.options.environment) {
        return 'Required option environment missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FLOW_REMOVE).helpInformation());
    log(
      `  Remarks:
  
    By default, the command will try to remove a Microsoft Flow you own.
    If you want to remove a Microsoft Flow owned by another user, use the 
    ${chalk.blue('asAdmin')} flag.

    If the environment with the name you specified doesn't exist, you will get
    the ${chalk.grey('Access to the environment \'xyz\' is denied.')} error.

    If the Microsoft Flow with the name you specified doesn't exist, you will
    get the ${chalk.grey(`Error: Resource \'abc\' does not exist in environment \'xyz\'`)}
    error.
   
  Examples:
  
    Removes the specified Microsoft Flow owned by the currently signed-in user
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d

    Removes the specified Microsoft Flow owned by the currently signed-in user
    without confirmation
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --confirm

    Removes the specified Microsoft Flow owned by another user
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin

    Removes the specified Microsoft Flow owned by another user without
    confirmation
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin --confirm
`);
  }
}

module.exports = new FlowRemoveCommand();