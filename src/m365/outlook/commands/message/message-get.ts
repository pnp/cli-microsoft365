import auth, { Auth } from '../../../../Auth';
import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  userId?: string;
  userPrincipalName?: string;
}

class OutlookMessageGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_GET;
  }

  public get description(): string {
    return 'Retrieves specified message';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userPrincipalName: typeof args.options.userPrincipalName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userPrincipalName [userPrincipalName]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAuth: boolean | undefined = Auth.isAppOnlyAuth(auth.service.accessTokens[this.resource].accessToken);
      if (this.verbose) {
        logger.logToStderr(`Retrieving message with id ${args.options.id} using ${isAppOnlyAuth ? 'app only permissions' : 'delegated permissions'}`);
      }

      let requestUrl = '';

      if (isAppOnlyAuth) {
        if (!args.options.userId && !args.options.userPrincipalName) {
          throw `The option 'userId' or 'userPrincipalName' is required when retrieving an email using app only credentials`;
        }
        if (args.options.userId && args.options.userPrincipalName) {
          throw `Both options 'userId' and 'userPrincipalName' cannot be set when retrieving an email using app only credentials`;
        }
        requestUrl += `users/${args.options.userId ? args.options.userId : args.options.userPrincipalName}`;
      }
      else {
        if (args.options.userId || args.options.userPrincipalName) {
          throw `Option 'userId' or 'userPrincipalName' is not allowed when using delegated credentials`;
        }
        requestUrl += 'me';
      }

      requestUrl += `/messages/${args.options.id}`;

      const requestOptions: AxiosRequestConfig = {
        url: `${this.resource}/v1.0/${requestUrl}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<any>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new OutlookMessageGetCommand();