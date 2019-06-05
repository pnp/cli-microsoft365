import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
  id: string;
}

class TeamsAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_APP_REMOVE}`;
  }

  public get description(): string {
    return 'Removes a Teams app from the organization\'s app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const { id: appId } = args.options;

    const removeApp: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        }
      };

      request
        .delete(requestOptions)
        .then((): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
    };

    if (args.options.confirm) {
      removeApp();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Teams app ${appId} from the app catalog?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeApp();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the Teams app to remove. Needs to be available in your organization\'s app catalog.'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the app'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    You can only remove a Teams app as a global administrator.

  Examples:

    Remove the Teams app with ID ${chalk.grey('83cece1e-938d-44a1-8b86-918cf6151957')} from
    the organization's app catalog. Will prompt for confirmation before actually
    removing the app.
      ${this.name} --id 83cece1e-938d-44a1-8b86-918cf6151957

    Remove the Teams app with ID ${chalk.grey('83cece1e-938d-44a1-8b86-918cf6151957')} from
    the organization's app catalog. Don't prompt for confirmation.
      ${this.name} --id 83cece1e-938d-44a1-8b86-918cf6151957 --confirm
`);
  }
}

module.exports = new TeamsAppRemoveCommand();