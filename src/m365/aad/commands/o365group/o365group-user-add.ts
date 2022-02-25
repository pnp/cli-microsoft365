import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import teamsCommands from '../../../teams/commands';
import commands from '../../commands';

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
    return commands.O365GROUP_USER_ADD;
  }

  public get description(): string {
    return 'Adds user to specified Microsoft 365 Group or Microsoft Teams team';
  }

  public alias(): string[] | undefined {
    return [teamsCommands.USER_ADD];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.role = args.options.role;
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    telemetryProps.groupId = typeof args.options.groupId !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const providedGroupId: string = (typeof args.options.groupId !== 'undefined') ? args.options.groupId : args.options.teamId as string;

    const requestOptions: any = {
      url: `${this.resource}/v1.0/users/${encodeURIComponent(args.options.userName)}/id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<{ value: string; }>(requestOptions)
      .then((res: { value: string; }): Promise<void> => {
        const endpoint: string = `${this.resource}/v1.0/groups/${providedGroupId}/${((typeof args.options.role !== 'undefined') ? args.options.role : '').toLowerCase() === 'owner' ? 'owners' : 'members'}/$ref`;

        const requestOptions: any = {
          url: endpoint,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: { "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + res.value }
        };

        return request.post(requestOptions);
      })
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --userName <userName>'
      },
      {
        option: "-i, --groupId [groupId]"
      },
      {
        option: "--teamId [teamId]"
      },
      {
        option: '-r, --role [role]',
        autocomplete: ['Owner', 'Member']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.groupId && !args.options.teamId) {
      return 'Please provide one of the following parameters: groupId or teamId';
    }

    if (args.options.groupId && args.options.teamId) {
      return 'You cannot provide both a groupId and teamId parameter, please provide only one';
    }

    if (args.options.teamId && !validation.isValidGuid(args.options.teamId as string)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
      return `${args.options.groupId} is not a valid GUID`;
    }

    if (args.options.role) {
      if (['owner', 'member'].indexOf(args.options.role.toLowerCase()) === -1) {
        return `${args.options.role} is not a valid role value. Allowed values Owner|Member`;
      }
    }

    return true;
  }
}

module.exports = new AadO365GroupUserAddCommand();