import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import * as fs from 'fs';
import * as path from 'path';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  filePath: string;
  overwrite?: boolean;
  scope?: string;
}

class SpoAppAddCommand extends SpoAppBaseCommand {
  public get name(): string {
    return `${commands.APP_ADD}`;
  }

  public get description(): string {
    return 'Adds an app to the specified SharePoint Online app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.overwrite = (!(!args.options.overwrite)).toString();
    telemetryProps.scope = args.options.scope || 'tenant';
    telemetryProps.appCatalogUrl = (!(!args.options.appCatalogUrl)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    const overwrite: boolean = args.options.overwrite || false;

    this
      .getSpoUrl(cmd, this.debug)
      .then((spoUrl: string): Promise<string> => {
        return this.getAppCatalogSiteUrl(cmd, spoUrl, args);
      })
      .then((appCatalogUrl: string): Promise<string> => {
        const fullPath: string = path.resolve(args.options.filePath);
        if (this.verbose) {
          cmd.log(`Adding app '${fullPath}' to app catalog...`);
        }

        const fileName: string = path.basename(fullPath);
        const requestOptions: any = {
          url: `${appCatalogUrl}/_api/web/${scope}appcatalog/Add(overwrite=${(overwrite.toString().toLowerCase())}, url='${fileName}')`,
          headers: {
            accept: 'application/json;odata=nometadata',
            binaryStringRequestBody: 'true'
          },
          body: fs.readFileSync(fullPath)
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: any = JSON.parse(res);
        if (args.options.output === 'json') {
          cmd.log(json);
        }
        else {
          cmd.log(json.UniqueId);
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --filePath <filePath>',
        description: 'Absolute or relative path to the solution package file to add to the app catalog'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the app catalog: tenant|sitecollection. Default tenant',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]',
        description: 'The URL of the app catalog where the solution package will be added. It must be specified when the scope is \'sitecollection\'',
      },
      {
        option: '--overwrite [overwrite]',
        description: 'Set to overwrite the existing package file'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      // verify either 'tenant' or 'sitecollection' specified if scope provided
      if (args.options.scope) {
        const testScope: string = args.options.scope.toLowerCase();
        if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
          return `Scope must be either 'tenant' or 'sitecollection'`;
        }

        if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
          return `You must specify appCatalogUrl when the scope is sitecollection`;
        }
      }

      if (!args.options.filePath) {
        return 'Missing required option filePath';
      }

      const fullPath: string = path.resolve(args.options.filePath);

      if (!fs.existsSync(fullPath)) {
        return `File '${fullPath}' not found`;
      }

      if (fs.lstatSync(fullPath).isDirectory()) {
        return `Path '${fullPath}' points to a directory`;
      }

      if (args.options.appCatalogUrl) {
        return SpoAppBaseCommand.isValidSharePointUrl(args.options.appCatalogUrl);
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (message: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_ADD).helpInformation());
    log(
      `  Remarks:

    When specifying the path to the app package file you can use both relative
    and absolute paths. Note, that ~ in the path, will not be resolved and will
    most likely result in an error.

    When adding an app to the tenant app catalog, it's not necessary to specify
    the tenant app catalog URL. When the URL is not specified, the CLI will
    try to resolve the URL itself. Specifying the app catalog URL is required
    when you want to add the app to a site collection app catalog.

    When specifying site collection app catalog, you can specify the URL either
    with our without the ${chalk.grey('AppCatalog')} part, for example
    ${chalk.grey('https://contoso.sharepoint.com/sites/team-a/AppCatalog')} or
    ${chalk.grey('https://contoso.sharepoint.com/sites/team-a')}. CLI will accept both formats.

    If you try to upload a package that already exists in the app catalog
    without specifying the ${chalk.blue('--overwrite')} option, the command will fail
    with an error stating that the specified package already exists.

  Examples:
  
    Add the ${chalk.grey('spfx.sppkg')} package to the tenant app catalog
      ${commands.APP_ADD} --filePath /Users/pnp/spfx/sharepoint/solution/spfx.sppkg

    Overwrite the ${chalk.grey('spfx.sppkg')} package in the tenant app catalog with the newer
    version
      ${commands.APP_ADD} --filePath sharepoint/solution/spfx.sppkg --overwrite

    Add the ${chalk.grey('spfx.sppkg')} package to the site collection app catalog 
    of site ${chalk.grey('https://contoso.sharepoint.com/sites/site1')}
      ${commands.APP_ADD} --filePath c:\\spfx.sppkg --scope sitecollection --appCatalogUrl https://contoso.sharepoint.com/sites/site1

  More information:

    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new SpoAppAddCommand();