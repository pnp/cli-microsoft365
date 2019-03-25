import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import GraphCommand from "../../GraphCommand";
import { Team } from './Team';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
}

class GraphTeamsMessagingSettingsListCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_MESSAGINGSETTINGS_LIST}`;
  }

  public get description(): string {
    return 'Lists messaging settings for a Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}?$select=messagingSettings`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: Team): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }
        
        cmd.log(res.messagingSettings);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team for which to get the messaging settings'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:

    To get messaging settings for a Microsoft Teams team, you have to first log
    in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:
         
    Get messaging settings for a Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${commands.TEAMS_MESSAGINGSETTINGS_LIST} --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
`);
  }
}

module.exports = new GraphTeamsMessagingSettingsListCommand();