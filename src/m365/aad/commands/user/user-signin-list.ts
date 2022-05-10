import { SignIn } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.userName = typeof args.options.userName !== 'undefined';
    telemetryProps.userId = typeof args.options.userId !== 'undefined';
    telemetryProps.appDisplayName = typeof args.options.appDisplayName !== 'undefined';
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'userPrincipalName', 'appId', 'appDisplayName', 'createdDateTime'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let endpoint: string = `${this.resource}/v1.0/auditLogs/signIns`;
    let filter: string = "";
    if (args.options.userName || args.options.userId) {
      filter = args.options.userId ?
        `?$filter=userId eq '${encodeURIComponent(args.options.userId as string)}'` :
        `?$filter=userPrincipalName eq '${encodeURIComponent(args.options.userName as string)}'`;
    }
    if (args.options.appId || args.options.appDisplayName) {
      filter += filter ? " and " : "?$filter=";
      filter += args.options.appId ?
        `appId eq '${encodeURIComponent(args.options.appId)}'` :
        `appDisplayName eq '${encodeURIComponent(args.options.appDisplayName as string)}'`;
    }
    endpoint += filter;
    odata
      .getAllItems<SignIn>(endpoint)
      .then((signins): void => {
        logger.log(signins);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
  public validate(args: CommandArgs): boolean | string {
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
}

module.exports = new AadUserSigninListCommand();
