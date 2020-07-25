import commands from '../../commands';
import teamsCommands from '../../../teams/commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import Utils from '../../../../Utils';
import { GroupUser } from './GroupUser';
import request from '../../../../request';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  role: string;
  teamId?: string;
  groupId?: string;
  userName: string;
}

class AadO365GroupUserSetCommand extends GraphItemsListCommand<GroupUser> {
  public get name(): string {
    return `${commands.O365GROUP_USER_SET}`;
  }

  public get description(): string {
    return 'Updates role of the specified user in the specified Microsoft 365 Group or Microsoft Teams team';
  }

  public alias(): string[] | undefined {
    return [teamsCommands.TEAMS_USER_SET];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    telemetryProps.groupId = typeof args.options.groupId !== 'undefined';
    telemetryProps.role = args.options.role;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const groupId: string = (typeof args.options.groupId !== 'undefined') ? args.options.groupId : args.options.teamId as string

    this
      .getOwners(cmd, groupId)
      .then((): Promise<void> => {
        return this.getMembersAndGuests(cmd, groupId);
      })
      .then((): Promise<void> | void => {
        // Filter out duplicate added values for owners (as they are returned as members as well)
        this.items = this.items.filter((groupUser, index, self) =>
          index === self.findIndex((t) => (
            t.id === groupUser.id && t.displayName === groupUser.displayName
          ))
        );

        if (this.debug) {
          cmd.log((typeof args.options.groupId !== 'undefined') ? 'Group owners and members:' : 'Team owners and members:');
          cmd.log(this.items);
          cmd.log('');
        }

        if (this.items.filter(i => i.userPrincipalName.toLocaleLowerCase() === args.options.userName.toLocaleLowerCase()).length <= 0) {
          const userNotInGroup = (typeof args.options.groupId !== 'undefined') ?
            'The specified user does not belong to the given Microsoft 365 Group. Please use the \'o365group user add\' command to add new users.' :
            'The specified user does not belong to the given Microsoft Teams team. Please use the \'graph teams user add\' command to add new users.';

          throw new Error(userNotInGroup);
        }

        if (args.options.role === "Owner") {
          const foundMember: GroupUser | undefined = this.items.find(e => e.userPrincipalName.toLocaleLowerCase() === args.options.userName.toLocaleLowerCase() && e.userType === 'Member');

          if (foundMember !== undefined) {
            const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners/$ref`;

            const requestOptions: any = {
              url: endpoint,
              headers: {
                'accept': 'application/json;odata.metadata=none'
              },
              json: true,
              body: { "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + foundMember.id }
            };

            return request.post(requestOptions);
          }
          else {
            const userAlreadyOwner = (typeof args.options.groupId !== 'undefined') ?
              'The specified user is already an owner in the specified Microsoft 365 group, and thus cannot be promoted.' :
              'The specified user is already an owner in the specified Microsoft Teams team, and thus cannot be promoted.';

            throw new Error(userAlreadyOwner);
          }
        }
        else {
          const foundOwner: GroupUser | undefined = this.items.find(e => e.userPrincipalName.toLocaleLowerCase() === args.options.userName.toLocaleLowerCase() && e.userType === 'Owner');

          if (foundOwner !== undefined) {
            const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners/${foundOwner.id}/$ref`;

            const requestOptions: any = {
              url: endpoint,
              headers: {
                'accept': 'application/json;odata.metadata=none'
              },
            };

            return request.delete(requestOptions);
          }
          else {
            const userAlreadyMember = (typeof args.options.groupId !== 'undefined') ?
              'The specified user is already a member in the specified Microsoft 365 group, and thus cannot be demoted.' :
              'The specified user is already a member in the specified Microsoft Teams team, and thus cannot be demoted.';

            throw new Error(userAlreadyMember);
          }
        }
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getOwners(cmd: CommandInstance, groupId: string): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners?$select=id,displayName,userPrincipalName,userType`;

    return this.getAllItems(endpoint, cmd, true).then((): void => {
      // Currently there is a bug in the Microsoft Graph that returns Owners as
      // userType 'member'. We therefore update all returned user as owner
      for (var i in this.items) {
        this.items[i].userType = 'Owner';
      }
    });
  }

  private getMembersAndGuests(cmd: CommandInstance, groupId: string): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,userType`;
    return this.getAllItems(endpoint, cmd, false);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "-i, --groupId [groupId]",
        description: "The ID of the Microsoft 365 group for which to update user"
      },
      {
        option: "--teamId [teamId]",
        description: "The ID of the Microsoft Teams team for which to update user"
      },
      {
        option: '-n, --userName <userName>',
        description: 'UPN of the user for whom to update the role (eg. johndoe@example.com)'
      },
      {
        option: '-r, --role <role>',
        description: 'Role to set for the given user in the specified Microsoft 365 Group or Microsoft Teams team. Allowed values: Owner|Member',
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

      if (['Owner', 'Member'].indexOf(args.options.role) === -1) {
        return `${args.options.role} is not a valid role value. Allowed values Owner|Member`;
      }

      return true;
    };
  }
}

module.exports = new AadO365GroupUserSetCommand();
