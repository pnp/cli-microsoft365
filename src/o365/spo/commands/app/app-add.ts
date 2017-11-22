import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import VerboseOption from '../../../../VerboseOption';
import * as request from 'request-promise-native';
import {
  CommandHelp,
  CommandOption,
  CommandValidate
} from '../../../../Command';
import { ContextInfo } from '../../spo';
import SpoCommand from '../../SpoCommand';
import * as fs from 'fs';
import * as fspath from 'path';
const vorpal: Vorpal = require('../../../../vorpal-init');


interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  path: string;
  overwrite?: boolean;
}

class AppAddCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_ADD;
  }

  public get description(): string {
    return 'Deploys/enables an app in the tenant app catalog';
  }


  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.File = args.options.path;
    telemetryProps.Overwrite = args.options.overwrite || false;
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const verbose: boolean = args.options.verbose || false;
    const overwrite: boolean = args.options.overwrite || false;
    const path: string = args.options.path || "";

    auth
      .ensureAccessToken(auth.service.resource, this, verbose)
      .then((accessToken: string): Promise<ContextInfo> => {
        if (verbose) {
          cmd.log(`Retrieved access token ${accessToken}`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/contextinfo`,
          headers: {
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        }

        if (verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(`Adding app...`);

        const data = fs.readFileSync(path);

        const parsedName = fspath.parse(path);
        const filename = `${parsedName.name}.${parsedName.ext}`

        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/tenantappcatalog/Add(overwrite=${overwrite}, url='${filename}')`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata=verbose',
            'X-RequestDigest': res.FormDigestValue,
            'binaryStringRequestBody': true,
          },
          body: data
        };

        if (verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);



      })
      .then((res: string): void => {

        if (verbose) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log("App deployed/enabled");

        cb();
      }, (err: any): void => {
        cmd.log(vorpal.chalk.red(`Error: ${err}`));
        cb();
      });
  }


  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --path <pathtofile>',
        description: 'The path to the app package to upload.'
      }
    ];

    const parentOptions: CommandOption[] | undefined = super.options();
    if (parentOptions) {
      return options.concat(parentOptions);
    }
    else {
      return options;
    }
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.path) {
        return false;
      }
      return true;
    };
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.APP_ADD).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
   
  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.APP_DEPLOY} --path 'pathtoyourfile.sppkg'
      Adds the specified app to the app catalog

`);
    };
  }
}

module.exports = new AppAddCommand();