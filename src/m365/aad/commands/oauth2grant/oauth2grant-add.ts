import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import AadCommand from '../../../base/AadCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  clientId: string;
  resourceId: string;
  scope: string;
}

class AadOAuth2GrantAddCommand extends AadCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_ADD;
  }

  public get description(): string {
    return 'Grant the specified service principal OAuth2 permissions to the specified resource';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Granting the service principal specified permissions...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/myorganization/oauth2PermissionGrants?api-version=1.6`,
      headers: {
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        "odata.type": "Microsoft.DirectoryServices.OAuth2PermissionGrant",
        "clientId": args.options.clientId,
        "consentType": "AllPrincipals",
        "principalId": null,
        "resourceId": args.options.resourceId,
        "scope": args.options.scope,
        "startTime": "0001-01-01T00:00:00",
        "expiryTime": "9000-01-01T00:00:00"
      }
    };

    request
      .post<void>(requestOptions)
      .then((): void => {
        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --clientId <clientId>'
      },
      {
        option: '-r, --resourceId <resourceId>'
      },
      {
        option: '-s, --scope <scope>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.clientId)) {
      return `${args.options.clientId} is not a valid GUID`;
    }

    if (!Utils.isValidGuid(args.options.resourceId)) {
      return `${args.options.resourceId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadOAuth2GrantAddCommand();