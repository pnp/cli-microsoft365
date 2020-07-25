import commands from '../../commands';
import teamsCommands from '../../../teams/commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName: string;
  role?: string;
  teamId?: string;
  groupId?: string;
}

class AadO365GroupUserAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.O365GROUP_USER_ADD}`;
  }

  public get description(): string {
    return 'Adds user to specified Microsoft 365 Group or Microsoft Teams team';
  }

  public alias(): string[] | undefined {
    return [teamsCommands.TEAMS_USER_ADD];
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

    const requestOptions: any = {
      url: `${this.resource}/v1.0/users/${encodeURIComponent(args.options.userName)}/id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    };

    request
      .get<{ value: string; }>(requestOptions)
      .then((res: { value: string; }): Promise<{}> => {
        const endpoint: string = `${this.resource}/v1.0/groups/${providedGroupId}/${((typeof args.options.role !== 'undefined') ? args.options.role : '').toLowerCase() === 'owner' ? 'owners' : 'members'}/$ref`;

        const requestOptions: any = {
          url: endpoint,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          },
          json: true,
          body: { "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + res.value }
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
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
        description: "The ID of the Microsoft 365 Group to which to add the user"
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

      if (args.options.role) {
        if (['owner', 'member'].indexOf(args.options.role.toLowerCase()) === -1) {
          return `${args.options.role} is not a valid role value. Allowed values Owner|Member`;
        }
      }

      return true;
    };
  }
}

module.exports = new AadO365GroupUserAddCommand();