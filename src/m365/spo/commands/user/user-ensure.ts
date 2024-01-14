import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  entraId?: string;
  aadId?: string;
  userName?: string;
}

class SpoUserEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.USER_ENSURE;
  }

  public get description(): string {
    return 'Ensures that a user is available on a specific site';
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
        entraId: typeof args.options.entraId !== 'undefined',
        aadId: typeof args.options.aadId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--entraId [entraId]'
      },
      {
        option: '--aadId [aadId]'
      },
      {
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.entraId && !validation.isValidGuid(args.options.entraId)) {
          return `${args.options.entraId} is not a valid GUID.`;
        }

        if (args.options.aadId && !validation.isValidGuid(args.options.aadId)) {
          return `${args.options.aadId} is not a valid GUID.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['entraId', 'aadId', 'userName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.aadId) {
      args.options.entraId = args.options.aadId;

      this.warn(logger, `Option 'aadId' is deprecated. Please use 'entraId' instead`);
    }

    if (this.verbose) {
      await logger.logToStderr(`Ensuring user ${args.options.entraId || args.options.userName} at site ${args.options.webUrl}`);
    }

    try {
      const requestBody = {
        logonName: args.options.userName || await this.getUpnByUserId(args.options.entraId!, logger)
      };

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/ensureuser`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUpnByUserId(entraId: string, logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving user principal name for user with id ${entraId}`);
    }

    return await entraUser.getUpnByUserId(entraId);
  }
}

export default new SpoUserEnsureCommand();