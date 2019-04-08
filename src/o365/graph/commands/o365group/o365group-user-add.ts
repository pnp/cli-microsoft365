import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import * as request from 'request-promise-native';
import GraphCommand from '../../GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName: string;
  role?: string;
  teamId?: string;
  groupId?: string;
}

class GraphO365GroupUserAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.O365GROUP_USER_ADD}`;
  }

  public get description(): string {
    return 'Adds user to specified Office 365 Group or Microsoft Teams team';
  }

  public alias(): string[] | undefined {
    return [commands.TEAMS_USER_ADD];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.role = args.options.role;
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    telemetryProps.groupId = typeof args.options.groupId !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const providedGroupId: string = (typeof args.options.groupId !== 'undefined') ? args.options.groupId : args.options.teamId as string

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/users/${encodeURIComponent(args.options.userName)}/id`,
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
      .then((res: { value: string; }): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        const endpoint: string = `${auth.service.resource}/v1.0/groups/${providedGroupId}/${((typeof args.options.role !== 'undefined') ? args.options.role : '').toLowerCase() === 'owner' ? 'owners' : 'members'}/$ref`;

        const requestOptions: any = {
          url: endpoint,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'accept': 'application/json;odata.metadata=none'
          }),
          json: true,
          body: { "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + res.value }
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
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
        option: '-n, --userName <userName>',
        description: 'User\'s UPN (user principal name, eg. johndoe@example.com)'
      },
      {
        option: "-i, --groupId [groupId]",
        description: "The ID of the Office 365 Group to which to add the user"
      },
      {
        option: "--teamId [teamId]",
        description: "The ID of the Teams team to which to add the user"
      },
      {
        option: '-r, --role [role]',
        description: 'The role to be assigned to the new user: Owner|Member. Default Member',
        autocomplete: ['Owner', 'Member']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.groupId && !args.options.teamId) {
        return 'Please provide one of the following parameters: groupId or teamId';
      }

      if (args.options.groupId && args.options.teamId) {
        return 'You cannot provide both a groupId and teamId parameter, please provide only one';
      }

      if (args.options.teamId && !Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (args.options.groupId && !Utils.isValidGuid(args.options.groupId as string)) {
        return `${args.options.groupId} is not a valid GUID`;
      }

      if (!args.options.userName) {
        return 'Required parameter userName missing';
      }

      if (args.options.role) {
        if (['owner', 'member'].indexOf(args.options.role.toLowerCase()) === -1) {
          return `${args.options.role} is not a valid role value. Allowed values Owner|Member`;
        }
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

    To add user to the specified Office 365 Group or Microsoft Teams team,
    you have to first log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)}
    command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    Add a new member to the specified Office 365 Group
      ${chalk.grey(config.delimiter)} ${this.name} --groupId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'

    Add a new owner to the specified Office 365 Group
      ${chalk.grey(config.delimiter)} ${this.name} --groupId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Owner

    Add a new member to the specified Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${(this.alias() as string[])[0]} --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'
`);
  }
}

module.exports = new GraphO365GroupUserAddCommand();