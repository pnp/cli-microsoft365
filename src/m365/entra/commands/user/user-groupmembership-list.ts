import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request,  { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { ODataResponse } from '../../../../utils/odata.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  userEmail?: string;
  securityEnabledOnly?: boolean;
}

class EntraUserGroupmembershipListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_GROUPMEMBERSHIP_LIST;
  }

  public get description(): string {
    return 'Retrieves all groups where the user is a member of';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        userEmail: typeof args.options.userEmail !== 'undefined',
        securityEnabledOnly: !!args.options.securityEnabledOnly
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --userId [userId]'
      },
      {
        option: '-n, --userName [userName]'
      },
      {
        option: '-e, --userEmail [userEmail]'
      },
      {
        option: '--securityEnabledOnly [securityEnabledOnly]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId as string)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName as string)) {
          return `${args.options.userName} is not a valid user principal name`;
        }

        if (args.options.userEmail && !validation.isValidUserPrincipalName(args.options.userEmail as string)) {
          return `${args.options.userEmail} is not a valid user email`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName', 'userEmail'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let userId = args.options.userId;

    try {
      if (args.options.userName) {
        userId = await entraUser.getUserIdByUpn(args.options.userName);
      }
      else if (args.options.userEmail) {
        userId = await entraUser.getUserIdByEmail(args.options.userEmail);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users/${userId}/getMemberGroups`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          securityEnabledOnly: !!args.options.securityEnabledOnly
        }
      };

      const results = await request.post<ODataResponse<string[]>>(requestOptions);
      await logger.log(results.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserGroupmembershipListCommand();