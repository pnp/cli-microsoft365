import * as request from 'request-promise-native';
import { RequestError } from 'request-promise-native/errors';
import auth, { Auth } from './Auth';
import { CommandInstance } from './cli';
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

  public head<TResponse>(options: request.OptionsWithUrl): Promise<TResponse> {
    options.method = 'HEAD';
    return this.execute(options);
  }

  private execute<TResponse>(options: request.OptionsWithUrl, resolve?: (res: TResponse) => void, reject?: (error: any) => void): Promise<TResponse> {
    if (!this._cmd) {
      return Promise.reject('Command reference not set on the request object');
    }

    return new Promise<TResponse>((_resolve: (res: TResponse) => void, _reject: (error: any) => void): void => {
      const resource: string = Auth.getResourceFromUrl(options.url.toString());

      ((): Promise<string> => {
        if (options.headers && options.headers['x-anonymous']) {
          return Promise.resolve('');
        }
        else {
          return auth.ensureAccessToken(resource, this._cmd as CommandInstance, request.debug)
        }
      })()
        .then((accessToken: string): Promise<TResponse> => {
          if (options.headers) {
            if (options.headers['x-anonymous']) {
              delete options.headers['x-anonymous'];
            }
            else {
              options.headers.authorization = `Bearer ${accessToken}`;
            }
          }
          return this.req(options);
        })
        .then((res: TResponse): void => {
          if (request.debug && res) {
            (this._cmd as CommandInstance).log('REQUEST response body');
            (this._cmd as CommandInstance).log(JSON.stringify(res));
          }
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