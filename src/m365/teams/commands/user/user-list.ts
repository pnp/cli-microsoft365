import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  role?: string;
  teamId: string;
}

class TeamsUserListCommand extends GraphCommand {
  private items: User[] = [];

  public get name(): string {
    return commands.USER_LIST;
  }

  public get description(): string {
    return "Lists users for the specified Microsoft Teams team";
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        role: args.options.role
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: "-i, --teamId <teamId>"
      },
      {
        option: "-r, --role [type]",
        autocomplete: ["Owner", "Member", "Guest"]
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId as string)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.role) {
          if (['Owner', 'Member', 'Guest'].indexOf(args.options.role) === -1) {
            return `${args.options.role} is not a valid role value. Allowed values Owner|Member|Guest`;
          }
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getOwners(logger, args.options.teamId)
      .then((): Promise<User[]> => {
        if (args.options.role === "Owner") {
          return Promise.resolve([]);
        }

        return this.getMembersAndGuests(logger, args.options.teamId);
      })
      .then((items): void => {
        this.items = this.items.concat(items);

        // Filter out duplicate added values for owners (as they are returned as members as well)
        // this aligns the output with what is displayed in the Teams UI
        this.items = this.items.filter((groupUser, index, self) =>
          index === self.findIndex((t) => (
            t.id === groupUser.id && t.displayName === groupUser.displayName
          ))
        );

        if (args.options.role) {
          this.items = this.items.filter(i => i.userType === args.options.role);
        }

        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getOwners(logger: Logger, groupId: string): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners?$select=id,displayName,userPrincipalName,userType`;

    return odata.getAllItems<User>(endpoint).then((items): void => {
      this.items = this.items.concat(items);

      // Currently there is a bug in the Microsoft Graph that returns Owners as
      // userType 'member'. We therefore update all returned user as owner
      for (const i in this.items) {
        this.items[i].userType = "Owner";
      }
    });
  }

  private getMembersAndGuests(logger: Logger, groupId: string): Promise<User[]> {
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,userType`;
    return odata.getAllItems(endpoint);
  }
}

module.exports = new TeamsUserListCommand();