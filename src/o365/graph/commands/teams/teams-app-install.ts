import * as request from 'request-promise-native';
import auth from '../../GraphAuth';
import Utils from '../../../../Utils';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId: string;
  teamId: string;
}

class GraphTeamsAppInstallCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_APP_INSTALL}`;
  }

  public get description(): string {
    return 'Installs an app from the catalog to a Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${auth.service.resource}/v1.0`

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${endpoint}/teams/${args.options.teamId}/installedApps`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'content-type': 'application/json;odata=nometadata',
            'accept': 'application/json;odata.metadata=none'
          }),
          json: true,
          body: {
            'teamsApp@odata.bind': `${endpoint}/appCatalogs/teamsApps/${args.options.appId}`
          }
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
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
        option: '--teamId <teamId>',
        description: 'The ID of the Microsoft Teams team to which to install the app'
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

      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To install an app to a Microsoft Teams team, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

    The ${chalk.grey(`appId`)} has to be the ID of the app from the Microsoft Teams App Catalog.
    Do not use the ID from the manifest of the zip app package.
    Use the ${chalk.blue(`graph teams app list`)} command to get this ID.

  Examples:

    Install an app from the catalog in a Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${this.name} --appId 4440558e-8c73-4597-abc7-3644a64c4bce --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
`);
  }
}

module.exports = new GraphTeamsAppInstallCommand();