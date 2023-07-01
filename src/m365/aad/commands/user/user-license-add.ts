import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId: string;
  userName: string;
  ids: string;
}

class AadUserLicenseAddCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_LICENSE_ADD;
  }

  public get description(): string {
    return 'Assigns a license to a user';
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
      },
      {
        option: '--ids <ids>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId as string)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.ids && args.options.ids.split(',').some(e => !validation.isValidGuid(e))) {
          return `${args.options.ids} contains one or more invalid GUIDs`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['userId', 'userName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const addLicenses = args.options.ids.split(',').map(x => { return { "disabledPlans": [], "skuId": x }; });
    const requestBody = { "addLicenses": addLicenses, "removeLicenses": [] };

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users/${args.options.userId || args.options.userName}/assignLicense`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadUserLicenseAddCommand();