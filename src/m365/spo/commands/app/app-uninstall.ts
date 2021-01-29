import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const uninstallApp: () => void = (): void => {
      const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';

      if (this.verbose) {
        logger.logToStderr(`Uninstalling app '${args.options.id}' from the site '${args.options.siteUrl}'...`);
      }

      const requestOptions: any = {
        url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/uninstall`,
        headers: {
          accept: 'application/json;odata=nometadata'
        }
      };

      request
        .post(requestOptions)
        .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      uninstallApp();
    }
    else {
      Cli.prompt({
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
        option: '-i, --id <id>'
      },
      {
        option: '-s, --siteUrl <siteUrl>'
      },
      {
        option: '--scope [scope]',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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
  }
}

module.exports = new SpoAppUninstallCommand();