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
  teamId?: string;
  groupId?: string;
  userName: string;
  confirm?: boolean;
}

class AadO365GroupUserRemoveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.O365GROUP_USER_REMOVE}`;
  }

  public get description(): string {
    return 'Removes the specified user from specified Microsoft 365 Group or Microsoft Teams team';
  }

  public alias(): string[] | undefined {
    return [teamsCommands.TEAMS_USER_REMOVE];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    telemetryProps.groupId = typeof args.options.groupId !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let userId = '';
    const groupId: string = (typeof args.options.groupId !== 'undefined') ? args.options.groupId : args.options.teamId as string

    const removeUser: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/users/${encodeURIComponent(args.options.userName)}/id`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
      };

      request
        .get<{ value: string; }>(requestOptions)
        .then((res: { value: string; }): Promise<any> => {
          userId = res.value;

          const requestOptions: any = {
            url: `${this.resource}/v1.0/groups/${groupId}/owners?$select=id,displayName,userPrincipalName,userType`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            json: true
          };

          return request.get(requestOptions);
        })
        .then((res: any): Promise<void> => {
          const userIsOwner: boolean = (res.value.filter((i: any) => i.userPrincipalName === args.options.userName).length > 0);
          const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/${userIsOwner ? 'owners' : 'members'}/${userId}/$ref`;

          const requestOptions: any = {
            url: endpoint,
            headers: {
              'accept': 'application/json;odata.metadata=none'
            },
          };

          return request.delete(requestOptions);
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeUser();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove ${args.options.userName} from the ${(typeof args.options.groupId !== 'undefined' ? 'group' : 'team')} ${groupId}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeUser();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "-i, --groupId [groupId]",
        description: "The ID of the Microsoft 365 Group from which to remove the user"
      },
      {
        option: "--teamId [teamId]",
        description: "The ID of the Microsoft Teams team from which to remove the user"
      },
      {
        option: '-n, --userName <userName>',
        description: 'User\'s UPN (user principal name), eg. johndoe@example.com'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing user from the specified Microsoft 365 Group or Microsoft Teams team'
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

      return true;
    };
  }
}

module.exports = new AadO365GroupUserRemoveCommand();