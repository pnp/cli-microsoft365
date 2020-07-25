import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import AadCommand from '../../../base/AadCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Granting the service principal specified permissions...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/myorganization/oauth2PermissionGrants?api-version=1.6`,
      headers: {
        'content-type': 'application/json'
      },
      json: true,
      body: {
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
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --clientId <clientId>',
        description: 'objectId of the service principal for which permissions should be granted'
      },
      {
        option: '-r, --resourceId <resourceId>',
        description: 'objectId of the AAD application to which permissions should be granted'
      },
      {
        option: '-s, --scope <scope>',
        description: 'Permissions to grant'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.clientId)) {
        return `${args.options.clientId} is not a valid GUID`;
      }

      if (!Utils.isValidGuid(args.options.resourceId)) {
        return `${args.options.resourceId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new AadOAuth2GrantAddCommand();