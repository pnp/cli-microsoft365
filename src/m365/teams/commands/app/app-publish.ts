import { AxiosRequestConfig } from 'axios';
import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
}

class TeamsAppPublishCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_PUBLISH;
  }

  public get description(): string {
    return 'Publishes Teams app to the organization\'s app catalog';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --filePath <filePath>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const fullPath: string = path.resolve(args.options.filePath);
    if (this.verbose) {
      logger.logToStderr(`Adding app '${fullPath}' to app catalog...`);
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/appCatalogs/teamsApps`,
      headers: {
        "content-type": "application/zip",
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: fs.readFileSync(fullPath)
    };

    try {
      const res = await request.post(requestOptions);

      logger.log(res);

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsAppPublishCommand();