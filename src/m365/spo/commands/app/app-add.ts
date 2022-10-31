import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  filePath: string;
  overwrite?: boolean;
  appCatalogScope?: string;
}

class SpoAppAddCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_ADD;
  }

  public get description(): string {
    return 'Adds an app to the specified SharePoint Online app catalog';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        overwrite: (!(!args.options.overwrite)).toString(),
        appCatalogScope: args.options.appCatalogScope || 'tenant',
        appCatalogUrl: (!(!args.options.appCatalogUrl)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --filePath <filePath>'
      },
      {
        option: '-s, --appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
      },
      {
        option: '--overwrite [overwrite]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        // verify either 'tenant' or 'sitecollection' specified if scope provided
        if (args.options.appCatalogScope) {
          const testScope: string = args.options.appCatalogScope.toLowerCase();
          if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
            return `appCatalogScope must be either 'tenant' or 'sitecollection'`;
          }

          if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
            return `You must specify appCatalogUrl when the appCatalogScope is sitecollection`;
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
          return validation.isValidSharePointUrl(args.options.appCatalogUrl);
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
    const overwrite: boolean = args.options.overwrite || false;

    try {
      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      const appCatalogUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

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

      const res = await request.post<string>(requestOptions);

      const json: { UniqueId: string; } = JSON.parse(res);
      if (args.options.output === 'json') {
        logger.log(json);
      }
      else {
        logger.log(json.UniqueId);
      }
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }
}

module.exports = new SpoAppAddCommand();