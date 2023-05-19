import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GraphCommand from '../../../base/GraphCommand';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { accessToken } from '../../../../utils/accessToken';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
}

class PurviewSensitivityLabelPolicySettingsListCommand extends GraphCommand {
  public get name(): string {
    return commands.SENSITIVITYLABEL_POLICYSETTINGS_LIST;
  }

  public get description(): string {
    return 'Get a list of policy settings for a sensitivity label';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
      this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
    }

    const requestUrl: string = args.options.userId || args.options.userName
      ? `${this.resource}/beta/users/${args.options.userId || args.options.userName}/security/informationProtection/labelPolicySettings`
      : `${this.resource}/beta/me/security/informationProtection/labelPolicySettings`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get<any>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PurviewSensitivityLabelPolicySettingsListCommand();