import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
  overwrite?: boolean; 
  scope?: string;
}

class SpoAppAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.APP_ADD}`;
  }

  public get description(): string {
    return 'Adds an app to the specified SharePoint Online app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.overwrite = args.options.overwrite || false;
    telemetryProps.scope = (!(!args.options.scope)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    const overwrite: boolean = args.options.overwrite || false;

    if (this.debug) {
      cmd.log(`Retrieving access token for ${auth.service.resource}...`);
    }

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((res: any): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const fullPath: string = path.resolve(args.options.filePath);
        if (this.verbose) {
          cmd.log(`Adding app '${fullPath}' to app catalog...`);
        }

        const fileName: string = path.basename(fullPath);
        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/${scope}appcatalog/Add(overwrite=${(overwrite.toString().toLowerCase())}, url='${fileName}')`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata=nometadata',
            binaryStringRequestBody: 'true'
          }),
          body: fs.readFileSync(fullPath)
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

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
        description: 'Specify the target app catalog: \'tenant\' or \'sitecollection\' (default = tenant)',
        autocomplete: ['tenant', 'sitecollection']
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
      // verify either 'tenant' or 'site' specified if scope provided
      if (args.options.scope) {
        const testScope: string = args.options.scope.toLowerCase();
        if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
          return `Scope must be either 'tenant' or 'sitecollection' if specified`
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

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (message: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_ADD).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
  using the ${chalk.blue(commands.LOGIN)} command.
                
  Remarks:

    To add an app to the tenant or site collection app catalog, you have to first log in to a SharePoint site using the
    ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    When specifying the path to the app package file you can use both relative and absolute paths.
    Note, that ~ in the path, will not be resolved and will most likely result in an error.

    If you try to upload a package that already exists in the app catalog without specifying
    the ${chalk.blue('--overwrite')} option, the command will fail with an error stating that the
    specified package already exists.

  Examples:
  
    Add the ${chalk.grey('spfx.sppkg')} package to the tenant app catalog
      ${chalk.grey(config.delimiter)} ${commands.APP_ADD} --filePath /Users/pnp/spfx/sharepoint/solution/spfx.sppkg

    Overwrite the ${chalk.grey('spfx.sppkg')} package in the tenant app catalog with the newer version
      ${chalk.grey(config.delimiter)} ${commands.APP_ADD} --filePath sharepoint/solution/spfx.sppkg --overwrite

    Add the ${chalk.grey('spfx.sppkg')} package to the site collection app catalog 
    of the site you are currently logged in
      ${chalk.grey(config.delimiter)} ${commands.APP_ADD} --filePath c:/spfx.sppkg --scope sitecollection

  More information:

    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new SpoAppAddCommand();