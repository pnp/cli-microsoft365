import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import teamsCommands from '../../../teams/commands.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface UserResponse {
  value: string
}

interface Options extends GlobalOptions {
  teamId?: string;
  groupId?: string;
  userName: string;
  force?: boolean;
}

class AadM365GroupUserRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_USER_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified user from specified Microsoft 365 Group or Microsoft Teams team';
  }

  public alias(): string[] | undefined {
    return [teamsCommands.USER_REMOVE];
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
        force: (!(!args.options.force)).toString(),
        teamId: typeof args.options.teamId !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined'
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
        option: '-f, --force'
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

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['groupId', 'teamId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const groupId: string = (typeof args.options.groupId !== 'undefined') ? args.options.groupId : args.options.teamId as string;

    const removeUser = async (): Promise<void> => {
      try {
        // retrieve user
        const user: UserResponse = await request.get({
          url: `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.userName)}/id`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        });

        // used to verify if the group exists or not
        await request.get({
          url: `${this.resource}/v1.0/groups/${groupId}/id`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          }
        });

        try {
          // try to delete the user from the owners. Accepted error is 404
          await request.delete({
            url: `${this.resource}/v1.0/groups/${groupId}/owners/${user.value}/$ref`,
            headers: {
              'accept': 'application/json;odata.metadata=none'
            }
          });
        }
        catch (err: any) {
          // the 404 error is accepted
          if (err.response.status !== 404) {
            throw err.response.data;
          }
        }

        // try to delete the user from the members. Accepted error is 404
        try {
          await request.delete({
            url: `${this.resource}/v1.0/groups/${groupId}/members/${user.value}/$ref`,
            headers: {
              'accept': 'application/json;odata.metadata=none'
            }
          });
        }
        catch (err: any) {
          // the 404 error is accepted
          if (err.response.status !== 404) {
            throw err.response.data;
          }
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeUser();
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to remove ${args.options.userName} from the ${(typeof args.options.groupId !== 'undefined' ? 'group' : 'team')} ${groupId}?`);

      if (result) {
        await removeUser();
      }
    }
  }
}

export default new AadM365GroupUserRemoveCommand();