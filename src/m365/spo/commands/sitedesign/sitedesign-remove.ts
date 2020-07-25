import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  confirm?: boolean;
}

class SpoSiteDesignRemoveCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_REMOVE}`;
  }

  public get description(): string {
    return 'Removes the specified site design';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = args.options.confirm || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeSiteDesign: () => void = (): void => {
      let spoUrl: string = '';

      this
        .getSpoUrl(cmd, this.debug)
        .then((_spoUrl: string): Promise<ContextInfo> => {
          spoUrl = _spoUrl;
          return this.getRequestDigest(spoUrl);
        })
        .then((res: ContextInfo): Promise<void> => {
          const requestOptions: any = {
            url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteDesign`,
            headers: {
              'X-RequestDigest': res.FormDigestValue,
              'content-type': 'application/json;charset=utf-8',
              accept: 'application/json;odata=nometadata'
            },
            body: { id: args.options.id },
            json: true
          };

          return request.post(requestOptions);
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeSiteDesign();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the site design ${args.options.id}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeSiteDesign();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'Site design ID'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the site design'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new SpoSiteDesignRemoveCommand();