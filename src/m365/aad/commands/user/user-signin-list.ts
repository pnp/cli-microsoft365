import { SignIn } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName?: string;
  userId?: string;
  appDisplayName?: string;
  appId?: string;
}

class AadUserSigninListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_SIGNIN_LIST;
  }

  public get description(): string {
    return 'Retrieves the Azure AD user sign-ins for the tenant';
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
        userName: typeof args.options.userName !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        appDisplayName: typeof args.options.appDisplayName !== 'undefined',
        appId: typeof args.options.appId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --userName [userName]'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--appDisplayName [appDisplayName]'
      },
      {
        option: '--appId [appId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && args.options.userName) {
          return 'Specify either userId or userName, but not both';
        }

        if (args.options.appId && args.options.appDisplayName) {
          return 'Specify either appId or appDisplayName, but not both';
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId as string)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.appId && !validation.isValidGuid(args.options.appId as string)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'userPrincipalName', 'appId', 'appDisplayName', 'createdDateTime'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let endpoint: string = `${this.resource}/v1.0/auditLogs/signIns`;
      let filter: string = "";
      if (args.options.userName || args.options.userId) {
        filter = args.options.userId ?
          `?$filter=userId eq '${formatting.encodeQueryParameter(args.options.userId as string)}'` :
          `?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(args.options.userName as string)}'`;
      }
      if (args.options.appId || args.options.appDisplayName) {
        filter += filter ? " and " : "?$filter=";
        filter += args.options.appId ?
          `appId eq '${formatting.encodeQueryParameter(args.options.appId)}'` :
          `appDisplayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName as string)}'`;
      }
      endpoint += filter;

      const signins = await odata.getAllItems<SignIn>(endpoint);
      logger.log(signins);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadUserSigninListCommand();
