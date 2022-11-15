import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  filePath: string;
}

class TeamsAppUpdateCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_UPDATE;
  }

  public get description(): string {
    return 'Updates Teams app in the organization\'s app catalog';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-p, --filePath <filePath>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        const fullPath: string = path.resolve(args.options.filePath);

        if (!fs.existsSync(fullPath)) {
          return `File '${fullPath}' not found`;
        }

        if (fs.lstatSync(fullPath).isDirectory()) {
          return `Path '${fullPath}' points to a directory`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'name']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { filePath } = args.options;

    try {
      const appId: string = await this.getAppId(args);
      const fullPath: string = path.resolve(filePath);
      if (this.verbose) {
        logger.logToStderr(`Updating app with id '${appId}' and file '${fullPath}' in the app catalog...`);
      }

      const requestOptions: any = {
        url: `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`,
        headers: {
          "content-type": "application/zip"
        },
        data: fs.readFileSync(fullPath)
      };

      await request.put(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getAppId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/appCatalogs/teamsApps?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { id: string; }[] }>(requestOptions)
      .then(response => {
        const app: { id: string; } | undefined = response.value[0];

        if (!app) {
          return Promise.reject(`The specified Teams app does not exist`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple Teams apps with name ${args.options.name} found. Please choose one of these ids: ${response.value.map(x => x.id).join(', ')}`);
        }

        return Promise.resolve(app.id);
      });
  }
}

module.exports = new TeamsAppUpdateCommand();