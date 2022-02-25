import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const { filePath } = args.options;

    this
      .getAppId(args)
      .then((appId: string): Promise<void> => {
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

        return request.put(requestOptions);
      })
      .then(_ => cb(), (res: any): void => this.handleRejectedODataJsonPromise(res, logger, cb));
  }

  private getAppId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/appCatalogs/teamsApps?$filter=displayName eq '${encodeURIComponent(args.options.name as string)}'`,
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-p, --filePath <filePath>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.id && !args.options.name) {
      return 'Specify either id or name';
    }

    if (args.options.id && args.options.name) {
      return 'Specify either id or name, but not both';
    }

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
}

module.exports = new TeamsAppUpdateCommand();