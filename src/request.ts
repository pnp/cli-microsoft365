import Axios, { AxiosError, AxiosInstance, AxiosPromise, AxiosRequestConfig, AxiosResponse } from 'axios';
import auth, { Auth } from './Auth';
import { CommandInstance } from './cli';
import Utils from './Utils';
const packageJSON = require('../package.json');

class Request {
  private req: AxiosInstance;
  private _cmd?: CommandInstance;
  private _debug: boolean = false;

  public set debug(debug: boolean) {
    this._debug = debug;

    if (this._debug) {
      this.req.interceptors.request.use((config: AxiosRequestConfig): AxiosRequestConfig => {
        (this._cmd as CommandInstance).log('Request:');
        (this._cmd as CommandInstance).log(Utils.filterObject(config, ['url', 'method', 'headers', 'data', 'responseType', 'decompress']));
        return config;
      });
      // since we're stubbing requests, response interceptor is never called in
      // tests, so let's exclude it from coverage
      /* c8 ignore next 5 */
      this.req.interceptors.response.use((response: AxiosResponse): AxiosResponse => {
        (this._cmd as CommandInstance).log('Response:');
        (this._cmd as CommandInstance).log(Utils.filterObject(response, ['data', 'status', 'statusText', 'headers']));
        return response;
      }, (error: AxiosError): void => {
        (this._cmd as CommandInstance).log('Request error');
        (this._cmd as CommandInstance).log(Utils.filterObject(error.response, ['data', 'status', 'statusText', 'headers']));
        throw error;
      });
    }
  }

  public set cmd(cmd: CommandInstance) {
    this._cmd = cmd;
  }

  constructor() {
    this.req = Axios.create({
      headers: {
        'user-agent': `NONISV|SharePointPnP|Office365CLI/${packageJSON.version}`,
        'accept-encoding': 'gzip, deflate'
      },
      decompress: true,
      responseType: 'text',
      /* c8 ignore next */
      transformResponse: [data => data]
    });
    // since we're stubbing requests, request interceptor is never called in
    // tests, so let's exclude it from coverage
    /* c8 ignore next 7 */
    this.req.interceptors.request.use((config: AxiosRequestConfig): AxiosRequestConfig => {
      if (config.responseType === 'json') {
        config.transformResponse = Axios.defaults.transformResponse;
      }

      return config;
    });
  }

  public post<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'POST';
    return this.execute(options);
  }

  public get<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'GET';
    return this.execute(options);
  }

  public patch<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'PATCH';
    return this.execute(options);
  }

  public put<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'PUT';
    return this.execute(options);
  }

  public delete<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'DELETE';
    return this.execute(options);
  }

  public head<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'HEAD';
    return this.execute(options);
  }

  private execute<TResponse>(options: AxiosRequestConfig, resolve?: (res: TResponse) => void, reject?: (error: any) => void): Promise<TResponse> {
    if (!this._cmd) {
      return Promise.reject('Command reference not set on the request object');
    }

    return new Promise<TResponse>((_resolve: (res: TResponse) => void, _reject: (error: any) => void): void => {
      const resource: string = Auth.getResourceFromUrl(options.url as string);

      ((): Promise<string> => {
        if (options.headers && options.headers['x-anonymous']) {
          return Promise.resolve('');
        }
        else {
          return auth.ensureAccessToken(resource, this._cmd as CommandInstance, this._debug)
        }
      })()
        .then((accessToken: string): AxiosPromise<TResponse> => {
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
        .then((res: AxiosResponse<TResponse>): void => {
          if (resolve) {
            resolve(res.data);
          }
          else {
            _resolve(res.data);
          }
        }, (error: AxiosError): void => {
          if (error && error.response &&
            (error.response.status === 429 ||
              error.response.status === 503)) {
            let retryAfter: number = parseInt(error.response.headers['retry-after'] || '10');
            if (isNaN(retryAfter)) {
              retryAfter = 10;
            }
            if (this._debug) {
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