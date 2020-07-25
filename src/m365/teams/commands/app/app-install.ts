import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId: string;
  teamId: string;
}

class TeamsAppInstallCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_APP_INSTALL}`;
  }

  public get description(): string {
    return 'Installs an app from the catalog to a Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0`

    const requestOptions: any = {
      url: `${endpoint}/teams/${args.options.teamId}/installedApps`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        'accept': 'application/json;odata.metadata=none'
      },
      json: true,
      body: {
        'teamsApp@odata.bind': `${endpoint}/appCatalogs/teamsApps/${args.options.appId}`
      }
    };

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--appId <appId>',
        description: 'The ID of the app to install'
      },
      {
        option: '--teamId <teamId>',
        description: 'The ID of the Microsoft Teams team to which to install the app'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.appId)) {
        return `${args.options.appId} is not a valid GUID`;
      }

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new TeamsAppInstallCommand();