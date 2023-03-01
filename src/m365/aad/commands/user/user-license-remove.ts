import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { Cli } from '../../../../cli/Cli';
import { AxiosRequestConfig } from 'axios';
import GraphCommand from '../../../base/GraphCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  ids: string;
  confirm?: boolean;
}

class AadUserLicenseRemoveCommand extends GraphCommand {

  public get name(): string {
    return commands.USER_LICENSE_REMOVE;
  }

  public get description(): string {
    return 'Removes a license from a user';
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
        userName: typeof args.options.userName !== 'undefined',
        confirm: !!args.options.confirm
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
      },
      {
        option: '--confirm'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['userId', 'userName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId as string)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid user principal name (UPN)`;
        }

        if (args.options.ids && args.options.ids.split(',').some(e => !validation.isValidGuid(e))) {
          return `${args.options.ids} contains one or more invalid GUIDs`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing the licenses for the user '${args.options.userId || args.options.userName}'...`);
    }

    if (args.options.confirm) {
      await this.deleteUserLicenses(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the licenses for the user '${args.options.userId || args.options.userName}'?`
      });

      if (result.continue) {
        await this.deleteUserLicenses(args);
      }
    }
  }

  private async deleteUserLicenses(args: CommandArgs): Promise<void> {
    const removeLicenses = args.options.ids.split(',');
    const requestBody = { "addLicenses": [], "removeLicenses": removeLicenses };

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/users/${args.options.userId || args.options.userName}/assignLicense`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadUserLicenseRemoveCommand();