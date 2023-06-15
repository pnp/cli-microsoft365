import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

class AadUserGetCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_GET;
  }

  public get description(): string {
    return 'Gets information about the specified user';
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
    const properties: string = args.options.properties ?
      `&$select=${args.options.properties.split(',').map(p => formatting.encodeQueryParameter(p.trim())).join(',')}` :
      '';

    let requestUrl: string = `${this.resource}/v1.0/users`;

    if (args.options.id) {
      requestUrl += `?$filter=id eq '${formatting.encodeQueryParameter(args.options.id as string)}'${properties}`;
    }
    else if (args.options.userName) {
      requestUrl += `?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(args.options.userName as string)}'${properties}`;
    }
    else if (args.options.email) {
      requestUrl += `?$filter=mail eq '${formatting.encodeQueryParameter(args.options.email as string)}'${properties}`;
    }

    if (args.options.withManager) {
      requestUrl += '&$expand=manager($select=displayName,userPrincipalName,id,mail)';
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<{ value: User[] }>(requestOptions);

      const identifier = args.options.id ? `id ${args.options.id}`
        : args.options.userName ? `user name ${args.options.userName}`
          : `email ${args.options.email}`;

      if (res.value.length === 0) {
        throw `The specified user with ${identifier} does not exist`;
      }

      if (res.value.length > 1) {
        throw `Multiple users with ${identifier} found. Please disambiguate (user names): ${res.value.map(a => a.userPrincipalName).join(', ')} or (ids): ${res.value.map(a => a.id).join(', ')}`;
      }

      logger.log(res.value[0]);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadUserGetCommand();
