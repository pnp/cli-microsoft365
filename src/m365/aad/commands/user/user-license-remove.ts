import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { Cli } from '../../../../cli/Cli.js';
import GraphCommand from '../../../base/GraphCommand.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  ids: string;
  force?: boolean;
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
        force: !!args.options.force
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
        option: '-f, --force'
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
      await logger.logToStderr(`Removing the licenses for the user '${args.options.userId || args.options.userName}'...`);
    }

    if (args.options.force) {
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

    const requestOptions: CliRequestOptions = {
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

export default new AadUserLicenseRemoveCommand();