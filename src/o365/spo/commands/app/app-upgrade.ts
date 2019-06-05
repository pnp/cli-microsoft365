import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  siteUrl: string;
  scope?: string;
}

class SpoAppUpgradeCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_UPGRADE;
  }

  public get description(): string {
    return 'Upgrades app in the specified site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'tenant';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';

    if (this.verbose) {
      cmd.log(`Upgrading app '${args.options.id}' in site '${args.options.siteUrl}'...`);
    }

    const requestOptions: any = {
      url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/upgrade`,
      headers: {
        accept: 'application/json;odata=nometadata'
      }
    };

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the app to upgrade'
      },
      {
        option: '-s, --siteUrl <siteUrl>',
        description: 'Absolute URL of the site to upgrade the app in'
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

      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (!args.options.siteUrl) {
        return 'Required parameter siteUrl missing';
      }

      return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_UPGRADE).helpInformation());
    log(
      `  Remarks:
  
    If the app with the specified ID doesn't exist in the app catalog,
    the command will fail with an error.
 
  Examples:
  
    Upgrade the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    in the ${chalk.grey('https://contoso.sharepoint.com')} site.
      ${commands.APP_UPGRADE} --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com

    Upgrade the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    in the ${chalk.grey('https://contoso.sharepoint.com')} site from site collection app catalog.
      ${commands.APP_INSTALL} --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com --scope sitecollection

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new SpoAppUpgradeCommand();