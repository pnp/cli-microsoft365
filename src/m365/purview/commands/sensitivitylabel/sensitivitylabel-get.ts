import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  userId?: string;
  userName?: string;
}

class PurviewSensitivityLabelGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SENSITIVITYLABEL_GET;
  }

  public get description(): string {
    return 'Retrieve the specified sensitivity label';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
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
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `'${args.options.id}' is not a valid GUID.`;
        }

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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
      this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving sensitivity label with id ${args.options.id}`);
    }

    const requestUrl: string = args.options.userId || args.options.userName
      ? `${this.resource}/beta/users/${args.options.userId || args.options.userName}/security/informationProtection/sensitivityLabels/${args.options.id}`
      : `${this.resource}/beta/me/security/informationProtection/sensitivityLabels/${args.options.id}`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewSensitivityLabelGetCommand();