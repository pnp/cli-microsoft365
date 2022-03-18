import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  role?: string;
  groupId: string;
}

class AadO365GroupUserListCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_USER_LIST;
  }

  public get description(): string {
    return "Lists users for the specified Microsoft 365 group";
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.role = args.options.role;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let users: User[] = [];

    this
      .getOwners(logger, args.options.groupId)
      .then((owners): Promise<User[]> => {
        users = owners;

        if (args.options.role === 'Owner') {
          return Promise.resolve([]);
        }

        return this.getMembersAndGuests(logger, args.options.groupId);
      })
      .then((membersAndGuests): void => {
        users = users.concat(membersAndGuests);

        if (args.options.role) {
          users = users.filter(i => i.userType === args.options.role);
        }

        logger.log(users);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getOwners(logger: Logger, groupId: string): Promise<User[]> {
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners?$select=id,displayName,userPrincipalName,userType`;

    return odata
      .getAllItems<User>(endpoint, logger)
      .then(users => {
        // Currently there is a bug in the Microsoft Graph that returns Owners as
        // userType 'member'. We therefore update all returned user as owner
        users.forEach(user => {
          user.userType = 'Owner';
        });

        return users;
      });
  }

  private getMembersAndGuests(logger: Logger, groupId: string): Promise<User[]> {
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,userType`;
    return odata.getAllItems<User>(endpoint, logger);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "-i, --groupId <groupId>"
      },
      {
        option: "-r, --role [type]",
        autocomplete: ["Owner", "Member", "Guest"]
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!validation.isValidGuid(args.options.groupId as string)) {
      return `${args.options.groupId} is not a valid GUID`;
    }

    if (args.options.role) {
      if (['Owner', 'Member', 'Guest'].indexOf(args.options.role) === -1) {
        return `${args.options.role} is not a valid role value. Allowed values Owner|Member|Guest`;
      }
    }

    return true;
  }
}

module.exports = new AadO365GroupUserListCommand();