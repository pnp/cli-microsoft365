import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

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
    return commands.APP_ADD;
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    const overwrite: boolean = args.options.overwrite || false;

    this
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<string> => {
        return this.getAppCatalogSiteUrl(logger, spoUrl, args);
      })
      .then((appCatalogUrl: string): Promise<string> => {
        const fullPath: string = path.resolve(args.options.filePath);
        if (this.verbose) {
          logger.logToStderr(`Adding app '${fullPath}' to app catalog...`);
        }

        const fileName: string = path.basename(fullPath);
        const requestOptions: any = {
          url: `${appCatalogUrl}/_api/web/${scope}appcatalog/Add(overwrite=${(overwrite.toString().toLowerCase())}, url='${fileName}')`,
          headers: {
            accept: 'application/json;odata=nometadata',
            binaryStringRequestBody: 'true'
          },
          data: fs.readFileSync(fullPath)
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: { UniqueId: string; } = JSON.parse(res);
        if (args.options.output === 'json') {
          logger.log(json);
        }
        else {
          logger.log(json.UniqueId);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --filePath <filePath>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
      },
      {
        option: '--overwrite [overwrite]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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
  }
}

module.exports = new SpoAppAddCommand();