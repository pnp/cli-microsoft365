import fs from 'fs';
import path from 'path';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';

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
  private readonly appCatalogScopeOptions = ['tenant', 'sitecollection'];

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
        overwrite: !!args.options.overwrite,
        appCatalogScope: args.options.appCatalogScope || 'tenant',
        appCatalogUrl: typeof args.options.appCatalogUrl !== 'undefined'
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
        autocomplete: this.appCatalogScopeOptions
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
        if (args.options.appCatalogScope) {
          const appCatalogScope = args.options.appCatalogScope.toLowerCase();
          if (this.appCatalogScopeOptions.indexOf(appCatalogScope) === -1) {
            return `${args.options.appCatalogScope} is not a valid appCatalogScope. Allowed values are: ${this.appCatalogScopeOptions.join(', ')}`;
          }

          if (appCatalogScope === 'sitecollection' && !args.options.appCatalogUrl) {
            return `You must specify appCatalogUrl when appCatalogScope is sitecollection`;
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
        await logger.logToStderr(`Adding app '${fullPath}' to app catalog...`);
      }

      const fileName: string = path.basename(fullPath);
      const requestOptions: CliRequestOptions = {
        url: `${appCatalogUrl}/_api/web/${scope}appcatalog/Add(overwrite=${(overwrite.toString().toLowerCase())}, url='${fileName}')`,
        headers: {
          accept: 'application/json;odata=nometadata',
          binaryStringRequestBody: 'true'
        },
        data: fs.readFileSync(fullPath)
      };

      const res = await request.post<string>(requestOptions);

      const json: { UniqueId: string; } = JSON.parse(res);
      if (!Cli.shouldTrimOutput(args.options.output)) {
        await logger.log(json);
      }
      else {
        await logger.log(json.UniqueId);
      }
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }
}

export default new SpoAppAddCommand();