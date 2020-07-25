import commands from '../../commands';
import teamsCommands from '../../../teams/commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import Utils from '../../../../Utils';
import { GroupUser } from './GroupUser';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  role?: string;
  teamId?: string;
  groupId?: string;
}

class AadO365GroupUserListCommand extends GraphItemsListCommand<GroupUser> {
  public get name(): string {
    return `${commands.O365GROUP_USER_LIST}`;
  }

  public get description(): string {
    return "Lists users for the specified Microsoft 365 group or Microsoft Teams team";
  }

  public alias(): string[] | undefined {
    return [teamsCommands.TEAMS_USER_LIST];
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

    this
      .getOwners(cmd, providedGroupId)
      .then((): Promise<void> => {
        if (args.options.role === "Owner") {
          return Promise.resolve();
        }

        return this.getMembersAndGuests(cmd, providedGroupId);
      })
      .then(
        (): void => {
          // Filter out duplicate added values for owners (as they are returned as members as well)
          this.items = this.items.filter((groupUser, index, self) =>
            index === self.findIndex((t) => (
              t.id === groupUser.id && t.displayName === groupUser.displayName
            ))
          );

          if (args.options.role) {
            this.items = this.items.filter(i => i.userType === args.options.role)
          }

          cmd.log(this.items);

          if (this.verbose) {
            cmd.log(chalk.green("DONE"));
          }

          cb();
        },
        (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb)
      );
  }

  private getOwners(cmd: CommandInstance, groupId: string): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners?$select=id,displayName,userPrincipalName,userType`;

    return this.getAllItems(endpoint, cmd, true).then(
      (): void => {
        // Currently there is a bug in the Microsoft Graph that returns Owners as
        // userType 'member'. We therefore update all returned user as owner
        for (var i in this.items) {
          this.items[i].userType = "Owner";
        }
      }
    );
  }

  private getMembersAndGuests(cmd: CommandInstance, groupId: string): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,userType`;
    return this.getAllItems(endpoint, cmd, false);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "-i, --groupId [groupId]",
        description: "The ID of the Microsoft 365 group for which to list users"
      },
      {
        option: "--teamId [teamId]",
        description: "The ID of the Microsoft Teams team for which to list users"
      },
      {
        option: "-r, --role [type]",
        description:
          "Filter the results to only users with the given role: Owner|Member|Guest",
        autocomplete: ["Owner", "Member", "Guest"]
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
        if (['Owner', 'Member', 'Guest'].indexOf(args.options.role) === -1) {
          return `${args.options.role} is not a valid role value. Allowed values Owner|Member|Guest`;
        }
      }

      return true;
    };
  }
}

module.exports = new AadO365GroupUserListCommand();