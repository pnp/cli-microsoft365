import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  siteUrl: string;
  confirm?: boolean;
  scope?: string;
}

class SpoAppUninstallCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_UNINSTALL;
  }

  public get description(): string {
    return 'Uninstalls an app from the site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    telemetryProps.scope = args.options.scope || 'tenant';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const uninstallApp: () => void = (): void => {
      const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';

      if (this.verbose) {
        cmd.log(`Uninstalling app '${args.options.id}' from the site '${args.options.siteUrl}'...`);
      }

      const requestOptions: any = {
        url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/uninstall`,
        headers: {
          accept: 'application/json;odata=nometadata'
        }
      };

      request
        .post(requestOptions)
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
    };

    if (args.options.confirm) {
      uninstallApp();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to uninstall the app ${args.options.id} from site ${args.options.siteUrl}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          uninstallApp();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the app to uninstall'
      },
      {
        option: '-s, --siteUrl <siteUrl>',
        description: 'Absolute URL of the site to uninstall the app from'
      },
      {
        option: '--scope [scope]',
        description: 'Scope of the app catalog: tenant|sitecollection. Default tenant',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming uninstalling the app'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.scope) {
        const testScope: string = args.options.scope.toLowerCase();
        if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
          return `Scope must be either 'tenant' or 'sitecollection' if specified`
        }
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
    };
  }
}

module.exports = new SpoAppUninstallCommand();