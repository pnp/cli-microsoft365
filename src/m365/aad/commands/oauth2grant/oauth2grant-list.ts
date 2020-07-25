import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import AadCommand from '../../../base/AadCommand';
import { OAuth2PermissionGrant } from './OAuth2PermissionGrant';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  clientId: string;
}

class AadOAuth2GrantListCommand extends AadCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_LIST;
  }

  public get description(): string {
    return 'Lists OAuth2 permission grants for the specified service principal';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving list of OAuth grants for the service principal...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/myorganization/oauth2PermissionGrants?api-version=1.6&$filter=clientId eq '${encodeURIComponent(args.options.clientId)}'`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<{ value: OAuth2PermissionGrant[] }>(requestOptions)
      .then((res: { value: OAuth2PermissionGrant[] }): void => {
        if (res.value && res.value.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(res.value);
          }
          else {
            cmd.log(res.value.map(g => {
              return {
                objectId: g.objectId,
                resourceId: g.resourceId,
                scope: g.scope
              };
            }));
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --clientId <clientId>',
        description: 'objectId of the service principal for which the configured OAuth2 permission grants should be retrieved'
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

      return true;
    };
  }
}

module.exports = new AadOAuth2GrantListCommand();