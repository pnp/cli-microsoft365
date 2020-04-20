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
  appId: string;
  userId: string;
}

class TeamsUserAppAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_USER_APP_ADD}`;
  }

  public get description(): string {
    return 'Install an app in the personal scope of the specified user';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/beta`

    const requestOptions: any = {
      url: `${endpoint}/users/${args.options.userId}/teamwork/installedApps`,
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
          cmd.log(vorpal.chalk.green('DONE'));
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
        option: '--userId <userId>',
        description: 'The ID of the user to install the app for'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.appId) {
        return 'Required parameter appId missing';
      }

      if (!Utils.isValidGuid(args.options.appId)) {
        return `${args.options.appId} is not a valid GUID`;
      }

      if (!args.options.userId) {
        return 'Required parameter userId missing';
      }

      if (!Utils.isValidGuid(args.options.userId)) {
        return `${args.options.userId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently in preview
    and is subject to change once the API reached general availability.

    The ${chalk.grey(`appId`)} has to be the ID of the app from the Microsoft Teams App Catalog.
    Do not use the ID from the manifest of the zip app package.
    Use the ${chalk.blue(`teams app list`)} command to get this ID.

  Examples:

    Install an app from the catalog for the specified user
      ${this.name} --appId 4440558e-8c73-4597-abc7-3644a64c4bce --userId 2609af39-7775-4f94-a3dc-0dd67657e900
`);
  }
}

module.exports = new TeamsUserAppAddCommand();