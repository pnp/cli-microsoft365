import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GraphCommand from '../../../base/GraphCommand';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import { accessToken } from '../../../../utils/accessToken';
import { odata } from '../../../../utils/odata';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
}

class PurviewSensitivityLabelListCommand extends GraphCommand {
  public get name(): string {
    return commands.SENSITIVITYLABEL_LIST;
  }

  public get description(): string {
    return 'Get a list of sensitivity labels';
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

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'isActive'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
      this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
    }

    let requestUrl: string = `${this.resource}/beta/`;
    if (args.options.userId || args.options.userName) {
      requestUrl += `users/${args.options.userId || args.options.userName}`;
    }
    else {
      requestUrl += 'me';
    }
    requestUrl += `/security/informationProtection/sensitivityLabels`;

    try {
      const items = await odata.getAllItems(requestUrl);
      logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PurviewSensitivityLabelListCommand();