import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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
        properties: args.options.properties
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

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'userName', 'email']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const properties: string = args.options.properties ?
      `&$select=${args.options.properties.split(',').map(p => encodeURIComponent(p.trim())).join(',')}` :
      '';

    let requestUrl: string = `${this.resource}/v1.0/users`;

    if (args.options.id) {
      requestUrl += `?$filter=id eq '${encodeURIComponent(args.options.id as string)}'${properties}`;
    }
    else if (args.options.userName) {
      requestUrl += `?$filter=userPrincipalName eq '${encodeURIComponent(args.options.userName as string)}'${properties}`;
    }
    else if (args.options.email) {
      requestUrl += `?$filter=mail eq '${encodeURIComponent(args.options.email as string)}'${properties}`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<{ value: User[] }>(requestOptions)
      .then((res: { value: User[] }): Promise<User> => {
        if (res.value.length === 1) {
          return Promise.resolve(res.value[0]);
        }

        const identifier = args.options.id ? `id ${args.options.id}`
          : args.options.userName ? `user name ${args.options.userName}`
            : `email ${args.options.email}`;

        if (res.value.length === 0) {
          return Promise.reject(`The specified user with ${identifier} does not exist`);
        }

        return Promise.reject(`Multiple users with ${identifier} found. Please disambiguate (user names): ${res.value.map(a => a.userPrincipalName).join(', ')} or (ids): ${res.value.map(a => a.id).join(', ')}`);
      })
      .then((res: User): void => {
        logger.log(res);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadUserGetCommand();
