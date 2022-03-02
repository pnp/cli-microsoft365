import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  objectId?: string;
  userPrincipalName?: string;
  accountEnabled?: boolean;
}

class AadUserSetCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_SET;
  }

  public get description(): string {
    return 'Updates information about the specified user';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.objectId = typeof args.options.objectId !== 'undefined';
    telemetryProps.userPrincipalName = typeof args.options.userPrincipalName !== 'undefined';
    telemetryProps.accountEnabled = (!!args.options.accountEnabled).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const manifest: any = this.mapRequestBody(args.options);

    const requestOptions: any = {
      url: `${this.resource}/v1.0/users/${encodeURIComponent(args.options.objectId ? args.options.objectId : args.options.userPrincipalName as string)}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json',
      data: manifest
    };

    request
      .patch(requestOptions)
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    const excludeOptions: string[] = [
      'debug',
      'verbose',
      'output',
      'objectId',
      'i',
      'userPrincipalName',
      'n',
      'accountEnabled'
    ];

    if (options.accountEnabled) {
      requestBody['AccountEnabled'] = String(options.accountEnabled) === "true";
    }

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        requestBody[key] = `${(<any>options)[key]}`;
      }
    });
    return requestBody;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --objectId [objectId]'
      },
      {
        option: '-n, --userPrincipalName [userPrincipalName]'
      },
      {
        option: '--accountEnabled [accountEnabled]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.objectId && !args.options.userPrincipalName) {
      return 'Specify either objectId or userPrincipalName';
    }

    if (args.options.objectId && args.options.userPrincipalName) {
      return 'Specify either objectId or userPrincipalName but not both';
    }

    if (args.options.objectId &&
      !validation.isValidGuid(args.options.objectId)) {
      return `${args.options.objectId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadUserSetCommand();
