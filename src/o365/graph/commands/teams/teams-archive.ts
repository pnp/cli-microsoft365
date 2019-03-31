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
  teamId: string;
  shouldSetSpoSiteReadOnlyForMembers?: false;
}

class GraphTeamsArchiveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_ARCHIVE}`;
  }

  public get description(): string {
    return 'Archive the specified Microsoft Team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${auth.service.resource}/v1.0`

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${endpoint}/teams/${args.options.teamId}/archive`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'content-type': 'application/json;odata=nometadata',
            'accept': 'application/json;odata.metadata=none'
          }),
          json: true,
          body: {
            'shouldSetSpoSiteReadOnlyForMembers': `${args.options.shouldSetSpoSiteReadOnlyForMembers ? args.options.shouldSetSpoSiteReadOnlyForMembers : false}`
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
        option: '--teamId <teamId>',
        description: 'The ID of the Microsoft Teams team to archive'
      },
      {
        option: '--shouldSetSpoSiteReadOnlyForMembers <true or false>',
        description: 'This optional parameter defines whether to set permissions for team members to read-only on the Sharepoint Online site associated with the team. Setting it to false or omitting the body altogether will result in this step being skipped.'
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

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:
    To archive a Microsoft Team, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.
    The ${chalk.grey(`teamId`)} has to be a valid GUID.
    The ${chalk.grey(`shouldSetSpoSiteReadOnlyForMembers`)} has to be true or false.
  Examples:
    Archive a Microsoft Team
      ${chalk.grey(config.delimiter)} ${this.name} --teamId f5dba91d-6494-4d5e-89a7-ad832f6946d6 --shouldSetSpoSiteReadOnlyForMembers false
`);
  }
}

module.exports = new GraphTeamsArchiveCommand();