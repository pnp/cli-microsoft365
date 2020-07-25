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
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
          cmd.log(chalk.green('DONE'));
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
}

module.exports = new SpoAppAddCommand();