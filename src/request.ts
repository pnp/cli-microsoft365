import * as request from 'request-promise-native';
import { RequestError } from 'request-promise-native/errors';
const packageJSON = require('../package.json');

class Request {
  private req: any;
  private _cmd?: CommandInstance;

  public set debug(debug: boolean) {
    (request as any).debug = debug;
  }

  public set cmd(cmd: CommandInstance) {
    this._cmd = cmd;
  }

  constructor() {
    this.req = request.defaults({
      headers: {
        'user-agent': `NONISV|SharePointPnP|Office365CLI/${packageJSON.version}`
      },
      gzip: true
    });
  }

  public post<TResponse>(options: request.OptionsWithUrl): Promise<TResponse> {
    options.method = 'POST';
    return this.execute(options);
  }

  public get<TResponse>(options: request.OptionsWithUrl): Promise<TResponse> {
    options.method = 'GET';
    return this.execute(options);
  }

  public patch<TResponse>(options: request.OptionsWithUrl): Promise<TResponse> {
    options.method = 'PATCH';
    return this.execute(options);
  }

  public put<TResponse>(options: request.OptionsWithUrl): Promise<TResponse> {
    options.method = 'PUT';
    return this.execute(options);
  }

  public delete<TResponse>(options: request.OptionsWithUrl): Promise<TResponse> {
    options.method = 'DELETE';
    return this.execute(options);
  }

  private execute<TResponse>(options: request.OptionsWithUrl, resolve?: (res: TResponse) => void, reject?: (error: any) => void): Promise<TResponse> {
    if (!this._cmd) {
      return Promise.reject('Command reference not set on the request object');
    }

    return new Promise<TResponse>((_resolve: (res: TResponse) => void, _reject: (error: any) => void): void => {
      this
        .req(options)
        .then((res: TResponse): void => {
          if (resolve) {
            resolve(res);
          }
          else {
            _resolve(res);
          }
        }, (error: RequestError): void => {
          if (error && error.response &&
            (error.response.statusCode === 429 ||
              error.response.statusCode === 503)) {
            let retryAfter: number = parseInt(error.response.headers['retry-after'] || '10');
            if (isNaN(retryAfter)) {
              retryAfter = 10;
            }
            if (request.debug) {
              (this._cmd as CommandInstance).log(`Request throttled. Waiting ${retryAfter}sec before retrying...`);
            }
            setTimeout(() => {
              this.execute(options, resolve || _resolve, reject || _reject);
            }, retryAfter * 1000);
          }
          else {
            if (reject) {
              reject(error);
            }
            else {
              _reject(error);
            }
          }
        });
    });
  }
}

export default new Request();