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
  scope?: string;
}

class SpoAppInstallCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_INSTALL;
  }

  public get description(): string {
    return 'Installs an app from the specified app catalog in the site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'tenant';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';

    if (this.verbose) {
      cmd.log(`Installing app '${args.options.id}' in site '${args.options.siteUrl}'...`);
    }

    const requestOptions: any = {
      url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/install`,
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
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the app to install'
      },
      {
        option: '-s, --siteUrl <siteUrl>',
        description: 'Absolute URL of the site to install the app in'
      },
      {
        option: '--scope [scope]',
        description: 'Scope of the app catalog: tenant|sitecollection. Default tenant',
        autocomplete: ['tenant', 'sitecollection']
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

module.exports = new SpoAppInstallCommand();