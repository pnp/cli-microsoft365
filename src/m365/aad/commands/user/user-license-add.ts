import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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
    return 'Assigns subscriptions to a user';
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

        if (!validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
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
    const requestBody = { "addLicenses": addLicenses };

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/users/${args.options.userId || args.options.userName}/assignLicense`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      const response = await request.post(requestOptions);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }
}

module.exports = new AadUserLicenseAddCommand();