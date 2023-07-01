import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

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
      const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
      if (this.verbose) {
        await logger.logToStderr(`Retrieving message with id ${args.options.id} using ${isAppOnlyAccessToken ? 'app only permissions' : 'delegated permissions'}`);
      }

      let requestUrl = '';

      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userPrincipalName) {
          throw `The option 'userId' or 'userPrincipalName' is required when retrieving an email using app only credentials`;
        }
        if (args.options.userId && args.options.userPrincipalName) {
          throw `Both options 'userId' and 'userPrincipalName' cannot be set when retrieving an email using app only credentials`;
        }
        requestUrl += `users/${args.options.userId ? args.options.userId : args.options.userPrincipalName}`;
      }
      else {
        if (args.options.userId && args.options.userPrincipalName) {
          throw `Both options 'userId' and 'userPrincipalName' cannot be set when retrieving an email using delegated credentials`;
        }

        if (args.options.userId || args.options.userPrincipalName) {
          requestUrl += `users/${args.options.userId ? args.options.userId : args.options.userPrincipalName}`;
        }
        else {
          requestUrl += 'me';
        }
      }

      requestUrl += `/messages/${args.options.id}`;

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/${requestUrl}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookMessageGetCommand();