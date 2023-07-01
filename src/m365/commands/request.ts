import { AxiosRequestConfig, AxiosResponse, RawAxiosRequestHeaders } from 'axios';
import fs from 'fs';
import path from 'path';
import { Logger } from '../../cli/Logger.js';
import Command from '../../Command.js';
import GlobalOptions from '../../GlobalOptions.js';
import request from '../../request.js';
import commands from './commands.js';
import auth from '../../Auth.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  method?: string;
  resource?: string;
  body?: string;
  filePath?: string;
}

class RequestCommand extends Command {
  private allowedMethods: string[] = ['get', 'post', 'put', 'patch', 'delete', 'head', 'options'];

  public get name(): string {
    return commands.REQUEST;
  }

  public get description(): string {
    return 'Executes the specified web request using CLI for Microsoft 365';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      const properties: any = {
        method: args.options.method || 'get',
        resource: typeof args.options.resource !== 'undefined',
        accept: args.options.accept || 'application/json',
        body: typeof args.options.body !== 'undefined',
        filePath: typeof args.options.filePath !== 'undefined'
      };

      const unknownOptions: any = this.getUnknownOptions(args.options);
      const unknownOptionsNames: string[] = Object.getOwnPropertyNames(unknownOptions);
      unknownOptionsNames.forEach(o => {
        properties[o] = typeof unknownOptions[o] !== 'undefined';
      });

      Object.assign(this.telemetryProperties, properties);
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '-m, --method [method]',
        autocomplete: this.allowedMethods
      },
      {
        option: '-r, --resource [resource]'
      },
      {
        option: '-b, --body [body]'
      },
      {
        option: '-p, --filePath [filePath]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const currentMethod = args.options.method || 'get';
        if (this.allowedMethods.indexOf(currentMethod) === -1) {
          return `${currentMethod} is not a valid value for method. Allowed values: ${this.allowedMethods.join(', ')}`;
        }

        if (args.options.body && (!args.options.method || args.options.method === 'get')) {
          return 'Specify a different method when using body';
        }

        if (args.options.body && !args.options['content-type']) {
          return 'Specify the content-type when using body';
        }

        if (args.options.filePath && !fs.existsSync(path.dirname(args.options.filePath))) {
          return 'The location specified in the filePath does not exist';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.debug) {
      await logger.logToStderr(`Preparing request...`);
    }

    try {
      const url = this.resolveUrlTokens(args.options.url);
      const method = (args.options.method || 'get').toUpperCase();
      const headers: RawAxiosRequestHeaders = {};

      const unknownOptions: any = this.getUnknownOptions(args.options);
      const unknownOptionsNames: string[] = Object.getOwnPropertyNames(unknownOptions);
      unknownOptionsNames.forEach(o => {
        headers[o] = unknownOptions[o];
      });

      if (!headers.accept) {
        headers.accept = 'application/json';
      }

      if (args.options.resource) {
        headers['x-resource'] = args.options.resource;
      }

      const config: AxiosRequestConfig<string> = {
        url: url,
        headers,
        method,
        data: args.options.body
      };

      if (headers.accept.toString().startsWith('application/json')) {
        config.responseType = 'json';
      }

      if (args.options.filePath) {
        config.responseType = 'stream';
      }

      if (this.verbose) {
        await logger.logToStderr(`Executing request...`);
      }


      if (args.options.filePath) {
        const file: AxiosResponse = await request.execute<AxiosResponse>(config);
        const filePath: string = await new Promise((resolve, reject) => {
          const writer = fs.createWriteStream(args.options.filePath as string);

          file.data.pipe(writer);

          writer.on('error', err => {
            reject(err);
          });
          writer.on('close', () => {
            resolve(args.options.filePath as string);
          });
        });

        if (this.verbose) {
          await logger.logToStderr(`File saved to path ${filePath}`);
        }
      }
      else {
        const res = await request.execute<string>(config);
        await logger.log(res);
      }
    }
    catch (err: any) {
      this.handleError(err);
    }
  }

  private resolveUrlTokens(url: string): string {
    if (url.startsWith('@graphbeta')) {
      return url.replace('@graphbeta', 'https://graph.microsoft.com/beta');
    }
    if (url.startsWith('@graph')) {
      return url.replace('@graph', 'https://graph.microsoft.com/v1.0');
    }
    if (url.startsWith('@spo')) {
      if (auth.service.spoUrl) {
        return url.replace('@spo', auth.service.spoUrl);
      }

      throw `SharePoint root site URL is unknown. Please set your SharePoint URL using command 'spo set'.`;
    }

    return url;
  }
}

export default new RequestCommand();