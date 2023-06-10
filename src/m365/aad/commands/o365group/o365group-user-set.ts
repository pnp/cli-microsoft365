import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import teamsCommands from '../../../teams/commands';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  role: string;
  teamId?: string;
  groupId?: string;
  userName: string;
}

class AadO365GroupUserSetCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_USER_SET;
  }

  public get description(): string {
    return 'Updates role of the specified user in the specified Microsoft 365 Group or Microsoft Teams team';
  }

  public alias(): string[] | undefined {
    return [teamsCommands.USER_SET];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        teamId: typeof args.options.teamId !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        role: args.options.role
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: "-i, --groupId [groupId]"
      },
      {
        option: "--teamId [teamId]"
      },
      {
        option: '-n, --userName <userName>'
      },
      {
        option: '-r, --role <role>',
        autocomplete: ['Owner', 'Member']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId as string)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (['Owner', 'Member'].indexOf(args.options.role) === -1) {
          return `${args.options.role} is not a valid role value. Allowed values Owner|Member`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['groupId', 'teamId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const groupId: string = (typeof args.options.groupId !== 'undefined') ? args.options.groupId : args.options.teamId as string;

      let users = await this.getOwners(groupId, logger);
      const membersAndGuests = await this.getMembersAndGuests(groupId, logger);
      users = users.concat(membersAndGuests);

      // Filter out duplicate added values for owners (as they are returned as members as well)
      users = users.filter((groupUser, index, self) =>
        index === self.findIndex((t) => (
          t.id === groupUser.id && t.displayName === groupUser.displayName
        ))
      );

      if (this.debug) {
        logger.logToStderr((typeof args.options.groupId !== 'undefined') ? 'Group owners and members:' : 'Team owners and members:');
        logger.logToStderr(users);
        logger.logToStderr('');
      }

      if (users.filter(i => args.options.userName.toUpperCase() === i.userPrincipalName!.toUpperCase()).length <= 0) {
        const userNotInGroup = (typeof args.options.groupId !== 'undefined') ?
          'The specified user does not belong to the given Microsoft 365 Group. Please use the \'o365group user add\' command to add new users.' :
          'The specified user does not belong to the given Microsoft Teams team. Please use the \'graph teams user add\' command to add new users.';

        throw new Error(userNotInGroup);
      }

      if (args.options.role === "Owner") {
        const foundMember: User | undefined = users.find(e => args.options.userName.toUpperCase() === e.userPrincipalName!.toUpperCase() && e.userType === 'Member');

        if (foundMember !== undefined) {
          const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners/$ref`;

          const requestOptions: CliRequestOptions = {
            url: endpoint,
            headers: {
              'accept': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: { "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + foundMember.id }
          };

          await request.post(requestOptions);
        }
        else {
          const userAlreadyOwner = (typeof args.options.groupId !== 'undefined') ?
            'The specified user is already an owner in the specified Microsoft 365 group, and thus cannot be promoted.' :
            'The specified user is already an owner in the specified Microsoft Teams team, and thus cannot be promoted.';

          throw new Error(userAlreadyOwner);
        }
      }
      else {
        const foundOwner: User | undefined = users.find(e => args.options.userName.toUpperCase() === e.userPrincipalName!.toUpperCase() && e.userType === 'Owner');

        if (foundOwner !== undefined) {
          const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners/${foundOwner.id}/$ref`;

          const requestOptions: CliRequestOptions = {
            url: endpoint,
            headers: {
              'accept': 'application/json;odata.metadata=none'
            }
          };

          await request.delete(requestOptions);
        }
        else {
          const userAlreadyMember = (typeof args.options.groupId !== 'undefined') ?
            'The specified user is already a member in the specified Microsoft 365 group, and thus cannot be demoted.' :
            'The specified user is already a member in the specified Microsoft Teams team, and thus cannot be demoted.';

          throw new Error(userAlreadyMember);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getOwners(groupId: string, logger: Logger): Promise<User[]> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving owners of the group with id ${groupId}`);
    }

    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/owners?$select=id,displayName,userPrincipalName,userType`;

    const users = await odata.getAllItems<User>(endpoint);
    // Currently there is a bug in the Microsoft Graph that returns Owners as
    // userType 'member'. We therefore update all returned user as owner
    users.forEach(user => {
      user.userType = 'Owner';
    });

    return users;

  }

  private async getMembersAndGuests(groupId: string, logger: Logger): Promise<User[]> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving members of the group with id ${groupId}`);
    }

    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,userType`;
    return await odata.getAllItems<User>(endpoint);
  }
}

module.exports = new AadO365GroupUserSetCommand();