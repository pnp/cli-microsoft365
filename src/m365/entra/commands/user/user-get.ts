import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  id?: string;
  userName?: string;
  email?: string;
  properties?: string;
  withManager?: boolean;
}

class EntraUserGetCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_GET;
  }

  public get description(): string {
    return 'Gets information about the specified user';
  }

  public alias(): string[] | undefined {
    return [aadCommands.USER_GET];
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
        id: typeof args.options.id !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        email: typeof args.options.email !== 'undefined',
        properties: args.options.properties,
        withManager: typeof args.options.withManager !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --userName [userName]'
      },
      {
        option: '--email [email]'
      },
      {
        option: '-p, --properties [properties]'
      },
      {
        option: '--withManager'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id &&
          !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'userName', 'email'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let userIdOrPrincipalName = args.options.id;

      if (args.options.userName) {
        // single user can be retrieved also by user principal name
        userIdOrPrincipalName = formatting.encodeQueryParameter(args.options.userName);
      }
      else if (args.options.email) {
        userIdOrPrincipalName = await entraUser.getUserIdByEmail(args.options.email);
      }

      const requestUrl: string = this.getRequestUrl(userIdOrPrincipalName!, args.options);

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const user = await request.get<User>(requestOptions);
      await logger.log(user);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestUrl(userIdOrPrincipalName: string, options: Options): string {
    const queryParameters: string[] = [];

    if (options.properties) {
      const allProperties = options.properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }
    }

    if (options.withManager) {
      queryParameters.push('$expand=manager($select=displayName,userPrincipalName,id,mail)');
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    // user principal name can start with $ but it violates the OData URL convention, so it must be enclosed in parenthesis and single quotes
    return userIdOrPrincipalName.startsWith('%24')
      ? `${this.resource}/v1.0/users('${userIdOrPrincipalName}')${queryString}`
      : `${this.resource}/v1.0/users/${userIdOrPrincipalName}${queryString}`;
  }
}

export default new EntraUserGetCommand();
