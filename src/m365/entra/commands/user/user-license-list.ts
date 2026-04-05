import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
}

class EntraUserLicenseListCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.USER_LICENSE_LIST;
  }

  public get description(): string {
    return 'Lists the license details for a given user';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'skuId', 'skuPartNumber'];
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
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid user principal name (UPN)`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['userId', 'userName'],
      runsWhen: (args) => args.options.userId || args.options.userName
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving licenses from user: ${args.options.userId || args.options.userName || 'current user'}.`);
    }

    let requestUrl: string = `${this.resource}/v1.0/`;
    if (args.options.userId || args.options.userName) {
      requestUrl += `users/${formatting.encodeQueryParameter(args.options.userId || args.options.userName as string)}`;
    }
    else {
      requestUrl += 'me';
    }
    requestUrl += '/licenseDetails';

    try {
      const items = await odata.getAllItems<any>(requestUrl);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserLicenseListCommand();