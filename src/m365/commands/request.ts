import { AxiosRequestConfig, AxiosRequestHeaders } from 'axios';
import auth, { Auth } from '../../Auth';
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
  headers?: string;
  accept?: string;
  contentType?: string;
}

class RequestCommand extends Command {
  public get name(): string {
    return commands.REQUEST;
  }

  public get description(): string {
    return 'Invoke a custom request at a Microsoft 365 API';
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
        method: args.options.method || 'get',
        resource: args.options.resource,
        accept: args.options.accept || 'application/json',
        contentType: args.options.contentType,
        body: typeof args.options.body !== 'undefined',
        headers: typeof args.options.headers !== 'undefined'
      });
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
        option: '-h, --headers [headers]'
      },
      {
        option: '-a, --accept [accept]'
      },
      {
        option: '-c, --contentType [contentType]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.body && (!args.options.method || args.options.method === "get")) {
          return "Specify a different method when using body";
        }

        if (args.options.body && !args.options.contentType) {
          return "Specify the contentType when using body";
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this.getRequestDetails(logger, args)
      .then(details => this.executeRequest(logger, details))
      .then(response => {
        if (response && response !== '') {
          if (!args.options.accept || args.options.accept.startsWith("application/json")) {
            logger.log(JSON.parse(response));
          }
          else {
            logger.log(response);
          }
        }
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private getRequestDetails(logger: Logger, args: CommandArgs): Promise<any> {
    if (this.verbose) {
      logger.logToStderr(`Preparing request...`);
    }

    const headers: AxiosRequestHeaders = args.options.headers ? JSON.parse(args.options.headers) : {};
    const method = (args.options.method || 'get').toUpperCase();

    if (args.options.accept || !headers['accept']) {
      headers['accept'] = args.options.accept || "application/json";
    }

    if (args.options.contentType) {
      headers['content-type'] = args.options.contentType;
    }

    const requestDetails: AxiosRequestConfig<string> = {
      url: args.options.url,
      headers,
      method,
      data: args.options.body
    };

    if (args.options.resource) {
      if (this.verbose) {
        logger.logToStderr(`Retrieving access token for resource...`);
      }

      const resource: string = Auth.getResourceFromUrl(args.options.resource);
      return auth
        .ensureAccessToken(resource, logger as Logger, this.debug)
        .then(accessToken => {
          requestDetails.headers!.authorization = `Bearer ${accessToken}`;

          return requestDetails;
        });
    }
    else {
      return Promise.resolve(requestDetails);
    }
  }

  private executeRequest(logger: Logger, options: AxiosRequestConfig): Promise<any> {
    if (this.verbose) {
      logger.logToStderr(`Executing request...`);
    }

    return request.execute(options);
  }
}

module.exports = new RequestCommand();