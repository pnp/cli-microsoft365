import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import request from '../../../../request';
import GraphCommand from '../../GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  allowCreateUpdateChannels?: string;
  allowDeleteChannels?: string;
  teamId: string;
}

class GraphTeamsGuestSettingsSetCommand extends GraphCommand {
  private static props: string[] = [
    'allowCreateUpdateChannels',
    'allowDeleteChannels'
  ];

  public get name(): string {
    return `${commands.TEAMS_GUESTSETTINGS_SET}`;
  }

  public get description(): string {
    return 'Updates guest settings of a Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    GraphTeamsGuestSettingsSetCommand.props.forEach(p => {
      telemetryProps[p] = (args.options as any)[p];
    });
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): Promise<{}> => {
        const body: any = {
          guestSettings: {}
        };
        GraphTeamsGuestSettingsSetCommand.props.forEach(p => {
          if (typeof (args.options as any)[p] !== 'undefined') {
            body.guestSettings[p] = (args.options as any)[p] === 'true';
          }
        });

        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none'
          },
          body: body,
          json: true
        };

        return request.patch(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Teams team for which to update settings'
      },
      {
        option: '--allowCreateUpdateChannels [allowCreateUpdateChannels]',
        description: 'Set to true to allow guests to create and update channels and to false to disallow it'
      },
      {
        option: '--allowDeleteChannels [allowDeleteChannels]',
        description: 'Set to true to allow guests to create and update channels and to false to disallow it'
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

      let isValid: boolean = true;
      let value, property: string = '';
      GraphTeamsGuestSettingsSetCommand.props.every(p => {
        property = p;
        value = (args.options as any)[p];
        isValid = typeof value === 'undefined' ||
          value === 'true' ||
          value === 'false';
        return isValid;
      });
      if (!isValid) {
        return `Value ${value} for option ${property} is not a valid boolean`;
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

    To update guest settings of the specified Microsoft Teams team, you have to
    first log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:
  
    Allow guests to create and edit channels
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --allowCreateUpdateChannels true

    Disallow guests to delete channels
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --allowDeleteChannels false
`);
  }
}

module.exports = new GraphTeamsGuestSettingsSetCommand();