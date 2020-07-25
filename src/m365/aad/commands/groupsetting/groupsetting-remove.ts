import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  confirm?: boolean;
}

class AadGroupSettingRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTING_REMOVE;
  }

  public get description(): string {
    return 'Removes the particular group setting';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeGroupSetting: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Removing group setting: ${args.options.id}...`);
      }

      const requestOptions: any = {
        url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
      };

      request
        .delete(requestOptions)
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
    };

    if (args.options.confirm) {
      removeGroupSetting();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the group setting ${args.options.id}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeGroupSetting();
        }
      });
    }
  }
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the group setting to remove'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the group setting'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new AadGroupSettingRemoveCommand();