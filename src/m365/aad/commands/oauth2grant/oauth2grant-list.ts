import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import AadCommand from '../../../base/AadCommand';
import commands from '../../commands';
import { OAuth2PermissionGrant } from './OAuth2PermissionGrant';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.log(`Retrieving list of OAuth grants for the service principal...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/myorganization/oauth2PermissionGrants?api-version=1.6&$filter=clientId eq '${encodeURIComponent(args.options.clientId)}'`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ value: OAuth2PermissionGrant[] }>(requestOptions)
      .then((res: { value: OAuth2PermissionGrant[] }): void => {
        if (res.value && res.value.length > 0) {
          if (args.options.output === 'json') {
            logger.log(res.value);
          }
          else {
            logger.log(res.value.map(g => {
              return {
                objectId: g.objectId,
                resourceId: g.resourceId,
                scope: g.scope
              };
            }));
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
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

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.clientId)) {
      return `${args.options.clientId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadOAuth2GrantListCommand();