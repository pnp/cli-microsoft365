import { AxiosRequestConfig, AxiosRequestHeaders } from 'axios';
import { Logger } from '../../cli';
import Command from '../../Command';
import GlobalOptions from '../../GlobalOptions';
import request from '../../request';
import commands from './commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  method?: string;
  resource?: string;
  body?: string;
  accept?: string;
}

class RequestCommand extends Command {
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
      const properties: any= {
        method: args.options.method || 'get',
        resource: typeof args.options.resource !== 'undefined',
        accept: args.options.accept || 'application/json',
        body: typeof args.options.body !== 'undefined'
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
        autocomplete: ['get', 'post', 'put', 'patch', 'delete', 'head', 'options']
      },
      {
        option: '-r, --resource [resource]'
      },
      {
        option: '-b, --body [body]'
      },
      {
        option: '-a, --accept [accept]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const allowedMethods = ['get', 'post', 'put', 'patch', 'delete', 'head', 'options'];
        const currentMethod = args.options.method || 'get';
        if (allowedMethods.indexOf(currentMethod) === -1) {
          return `${currentMethod} is not a valid value for method. Allowed values: ${allowedMethods.join(', ')}`;
        }

        if (args.options.body && (!args.options.method || args.options.method === 'get')) {
          return 'Specify a different method when using body';
        }

        if (args.options.body && !args.options['content-type']) {
          return 'Specify the contentType when using body';
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Preparing request...`);
    }

    const method = (args.options.method || 'get').toUpperCase();
    const headers: AxiosRequestHeaders = {
      accept: args.options.accept || 'application/json'
    };
      
    const unknownOptions: any = this.getUnknownOptions(args.options);
    const unknownOptionsNames: string[] = Object.getOwnPropertyNames(unknownOptions);
    unknownOptionsNames.forEach(o => {
      headers[o] = unknownOptions[o];
    });

    if (args.options.resource) {
      headers['x-resource'] = args.options.resource;
    }
    
    const config: AxiosRequestConfig<string> = {
      url: args.options.url,
      headers,
      method,
      data: args.options.body
    };

    if (headers.accept.toString().startsWith('application/json')) {
      config.responseType = 'json';
    }
    
    if (this.verbose) {
      logger.logToStderr(`Executing request...`);
    }
    
    request.execute<string>(config)
      .then(response => {
        logger.log(response);
        cb();
      }, (rawRes: any): void => this.handleError(rawRes, logger, cb));
  }

}

module.exports = new RequestCommand();