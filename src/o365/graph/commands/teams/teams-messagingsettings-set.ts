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
  teamId: string;
  allowUserEditMessages?: string;
  allowUserDeleteMessages?: string;
  allowOwnerDeleteMessages?: string;
  allowTeamMentions?: string;
  allowChannelMentions?: string;
}

class GraphTeamsMessageSettingsSetCommand extends GraphCommand {
  private static props: string[] = [
    'allowUserEditMessages',
    'allowUserDeleteMessages',
    'allowOwnerDeleteMessages',
    'allowTeamMentions',
    'allowChannelMentions'
  ];

  public get name(): string {
    return `${commands.TEAMS_MESSAGINGSETTINGS_SET}`;
  }

  public get description(): string {
    return 'Updates messaging settings of a Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    GraphTeamsMessageSettingsSetCommand.props.forEach((p: string) => {
        telemetryProps[p] = typeof (args.options as any)[p] !== 'undefined';
    });
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): Promise<{}> => {
        const body: any = {
          messagingSettings: {}
        };
        GraphTeamsMessageSettingsSetCommand.props.forEach((p: string) => {
          if (typeof (args.options as any)[p] !== 'undefined') {
            body.messagingSettings[p] = (args.options as any)[p].toLowerCase() === 'true';
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
        description: 'The ID of the Microsoft Teams team for which to update messaging settings'
      },
      {
        option: '--allowUserEditMessages [allowUserEditMessages]',
        description: 'Set to true to allow users to edit messages and to false to disallow it'
      },
      {
        option: '--allowUserDeleteMessages [allowUserDeleteMessages]',
        description: 'Set to true to allow users to delete messages and to false to disallow it'
      },
      {
        option: '--allowOwnerDeleteMessages [allowOwnerDeleteMessages]',
        description: 'Set to true to allow owner to delete messages and to false to disallow it'
      },
      {
        option: '--allowTeamMentions [allowTeamMentions]',
        description: 'Set to true to allow @team or @[team name] mentions and to false to disallow it'
      },
      {
        option: '--allowChannelMentions [allowChannelMentions]',
        description: 'Set to true to allow @channel or @[channel name] mentions and to false to disallow it'
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

      let hasDoublicate: boolean = false;
      let property: string = '';
      GraphTeamsMessageSettingsSetCommand.props.forEach((prop: string) => {
        if((args.options as any)[prop] instanceof Array) {
          property = prop;
          hasDoublicate = true;
        }
      });
      if(hasDoublicate) {
        return `Doublicated option ${property} specified. Specify only one`;
      }

      let isValid: boolean = true;
      let value: string = '';
      GraphTeamsMessageSettingsSetCommand.props.every((p: string) => {
        property = p;
        value = (args.options as any)[p];
        isValid = typeof value === 'undefined' || Utils.isValidBoolean(value)
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

    To update messaging settings of the specified Microsoft Teams team, you have to
    first log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:
  
    Allow users to edit messages in channels
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --allowUserEditMessages true

    Disallow users to delete messages in channels
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --allowUserDeleteMessages false
`);
  }
}

module.exports = new GraphTeamsMessageSettingsSetCommand();