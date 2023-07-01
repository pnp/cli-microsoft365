import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName: string;
  apiKey: string;
  domain?: string
}

class AadUserHibpCommand extends AnonymousCommand {
  public get name(): string {
    return commands.USER_HIBP;
  }

  public get description(): string {
    return 'Allows you to retrieve all accounts that have been pwned with the specified username';
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
        domain: args.options.domain
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --userName <userName>'
      },
      {
        option: '--apiKey, <apiKey>'
      },
      {
        option: '--domain, [domain]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidUserPrincipalName(args.options.userName)) {
          return 'Specify valid userName';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `https://haveibeenpwned.com/api/v3/breachedaccount/${formatting.encodeQueryParameter(args.options.userName)}${(args.options.domain ? `?domain=${formatting.encodeQueryParameter(args.options.domain)}` : '')}`,
        headers: {
          'accept': 'application/json',
          'hibp-api-key': args.options.apiKey,
          'x-anonymous': true
        },
        responseType: 'json'
      };

      const res = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      if ((err && err.response !== undefined && err.response.status === 404) && (this.debug || this.verbose)) {
        await logger.log('No pwnage found');
        return;
      }
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadUserHibpCommand();