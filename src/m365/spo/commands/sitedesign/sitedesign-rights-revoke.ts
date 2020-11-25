import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ContextInfo } from '../../spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  principals: string;
  confirm?: boolean;
}

class SpoSiteDesignRightsRevokeCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_RIGHTS_REVOKE}`;
  }

  public get description(): string {
    return 'Revokes access from a site design for one or more principals';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = args.options.confirm || false;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const revokePermissions: () => void = (): void => {
      let spoUrl: string = '';

      this
        .getSpoUrl(logger, this.debug)
        .then((_spoUrl: string): Promise<ContextInfo> => {
          spoUrl = _spoUrl;
          return this.getRequestDigest(spoUrl);
        })
        .then((res: ContextInfo): Promise<void> => {
          const requestOptions: any = {
            url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`,
            headers: {
              'X-RequestDigest': res.FormDigestValue,
              'content-type': 'application/json;charset=utf-8',
              accept: 'application/json;odata=nometadata'
            },
            data: {
              id: args.options.id,
              principalNames: args.options.principals.split(',').map(p => p.trim())
            },
            responseType: 'json'
          };

          return request.post(requestOptions);
        })
        .then((): void => {
          if (this.verbose) {
            logger.logToStderr(chalk.green('DONE'));
          }

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      revokePermissions();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to revoke access to site design ${args.options.id} from the specified users?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          revokePermissions();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the site design to revoke rights from'
      },
      {
        option: '-p, --principals <principals>',
        description: 'Comma-separated list of principals to revoke view rights from. Principals can be users or mail-enabled security groups in the form of "alias" or "alias@<domain name>.com"'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the site design'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoSiteDesignRightsRevokeCommand();
