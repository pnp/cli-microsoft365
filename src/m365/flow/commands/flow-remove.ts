import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../Command';
import request from '../../../request';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import Utils from '../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../cli';

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
            cmd.log(chalk.red(`Error: Resource '${args.options.name}' does not exist in environment '${args.options.environment}'`));
            cb();
          }
          else {
            if (this.verbose) {
              cmd.log(chalk.green('DONE'));
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
      if (!Utils.isValidGuid(args.options.name)) {
        return `${args.options.name} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new FlowRemoveCommand();