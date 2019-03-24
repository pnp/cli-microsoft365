import auth from '../../GraphAuth';
import Utils from '../../../../Utils';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../GraphCommand';
import request from '../../../../request';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
}

class GraphTeamsFunSettingsListCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_FUNSETTINGS_LIST}`;
  }

  public get description(): string {
    return 'Lists fun settings for the specified Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): Promise<{ funSettings: any; }> => {
        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}?$select=funSettings`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res: { funSettings: any }): void => {
        cmd.log(res.funSettings);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  };


  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team for which to list fun settings'
      },
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

    To get fun settings of a Microsoft Teams team, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    List fun settings of a Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${this.name} --teamId 83cece1e-938d-44a1-8b86-918cf6151957
`);
  }
}

module.exports = new GraphTeamsFunSettingsListCommand();