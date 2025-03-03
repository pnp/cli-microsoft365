import Axios, { AxiosError, AxiosInstance, AxiosProxyConfig, AxiosRequestConfig, AxiosResponse } from 'axios';
import { Stream } from 'stream';
import auth, { Auth, CloudType } from './Auth.js';
import { Logger } from './cli/Logger.js';
import { app } from './utils/app.js';
import { formatting } from './utils/formatting.js';
import { timings } from './cli/timings.js';
import { timersUtil } from './utils/timersUtil.js';

export interface CliRequestOptions extends AxiosRequestConfig {
  fullResponse?: boolean;
}

class Request {
  private req: AxiosInstance;
  private _logger?: Logger;
  private _debug: boolean = false;

  public set debug(debug: boolean) {
    // if the value to set is the same as current value return early to avoid
    // instantiating interceptors multiple times. This can happen when calling
    // one command from another
    if (this._debug === debug) {
      return;
    }

    this._debug = debug;

    if (this._debug) {
      this.req.interceptors.request.use(async config => {
        if (this._logger) {
          await this._logger.logToStderr('Request:');
          const properties: string[] = ['url', 'method', 'headers', 'responseType', 'decompress'];
          if (config.responseType !== 'stream') {
            properties.push('data');
          }
          await this._logger.logToStderr(JSON.stringify(formatting.filterObject(config, properties), null, 2));
        }
        return config;
      });
      // since we're stubbing requests, response interceptor is never called in
      // tests, so let's exclude it from coverage
      /* c8 ignore next 26 */
      this.req.interceptors.response.use(async (response: AxiosResponse): Promise<AxiosResponse> => {
        if (this._logger) {
          await this._logger.logToStderr('Response:');
          const properties: string[] = ['status', 'statusText', 'headers'];
          if (response.headers['content-type'] &&
            response.headers['content-type'].indexOf('json') > -1) {
            properties.push('data');
          }
          await this._logger.logToStderr(JSON.stringify({
            url: response.config.url,
            ...formatting.filterObject(response, properties)
          }, null, 2));
        }
        return response;
      }, async (error: AxiosError): Promise<void> => {
        if (this._logger) {
          const properties: string[] = ['status', 'statusText', 'headers'];
          await this._logger.logToStderr('Request error:');
          await this._logger.logToStderr(JSON.stringify({
            url: error.config?.url,
            ...formatting.filterObject(error.response, properties),
            error: (error as any).error
          }, null, 2));
        }
        throw error;
      });
    }
  }

  public get logger(): Logger | undefined {
    return this._logger;
  }

  public set logger(logger: Logger | undefined) {
    this._logger = logger;
  }

  constructor() {
    this.req = Axios.create({
      headers: {
        'user-agent': `NONISV|SharePointPnP|CLIMicrosoft365/${app.packageJson().version}`,
        'accept-encoding': 'gzip, deflate',
        'X-ClientService-ClientTag': `M365CLI:${app.packageJson().version}`
      },
      decompress: true,
      responseType: 'text',
      /* c8 ignore next */
      transformResponse: [data => data],
      maxBodyLength: Infinity,
      maxContentLength: Infinity
    });
    // since we're stubbing requests, request interceptor is never called in
    // tests, so let's exclude it from coverage
    /* c8 ignore next 7 */
    this.req.interceptors.request.use(config => {
      if (config.responseType === 'json') {
        config.transformResponse = Axios.defaults.transformResponse;
      }

      return config;
    });
    // since we're stubbing requests, response interceptor is never called in
    // tests, so let's exclude it from coverage
    /* c8 ignore next 15 */
    this.req.interceptors.response.use(
      (response: AxiosResponse) => response,
      (error: AxiosError<any>): void => {
        if (error &&
          error.response &&
          error.response.data &&
          !(error.response.data instanceof Stream)) {
          // move error details from response.data to error property to make
          // it compatible with our code
          (error as any).error = JSON.parse(JSON.stringify(error.response.data));
        }

        throw error;
      }
    );
  }

  public post<TResponse>(options: CliRequestOptions): Promise<TResponse> {
    options.method = 'POST';
    return this.execute(options);
  }

  public get<TResponse>(options: CliRequestOptions): Promise<TResponse> {
    options.method = 'GET';
    return this.execute(options);
  }

  public patch<TResponse>(options: CliRequestOptions): Promise<TResponse> {
    options.method = 'PATCH';
    return this.execute(options);
  }

  public put<TResponse>(options: CliRequestOptions): Promise<TResponse> {
    options.method = 'PUT';
    return this.execute(options);
  }

  public delete<TResponse>(options: CliRequestOptions): Promise<TResponse> {
    options.method = 'DELETE';
    return this.execute(options);
  }

  public head<TResponse>(options: CliRequestOptions): Promise<TResponse> {
    options.method = 'HEAD';
    return this.execute(options);
  }

  public async execute<TResponse>(options: CliRequestOptions): Promise<TResponse> {
    const start = process.hrtime.bigint();

    if (!this._logger) {
      throw 'Logger not set on the request object';
    }

    this.updateRequestForCloudType(options, auth.connection.cloudType);
    this.removeDoubleSlashes(options);

    try {
      let accessToken = '';

      if (options.headers && options.headers['x-anonymous']) {
        accessToken = '';
      }
      else {
        const url = options.headers && options.headers['x-resource'] ? options.headers['x-resource'] : options.url;
        const resource: string = Auth.getResourceFromUrl(url as string);
        accessToken = await auth.ensureAccessToken(resource, this._logger as Logger, this._debug);
      }

      if (options.headers) {
        if (options.headers['x-anonymous']) {
          delete options.headers['x-anonymous'];
        }
        if (options.headers['x-resource']) {
          delete options.headers['x-resource'];
        }
        if (accessToken !== '') {
          options.headers.authorization = `Bearer ${accessToken}`;
        }
      }

      const proxyUrl = process.env.HTTP_PROXY || process.env.HTTPS_PROXY;

      if (proxyUrl) {
        options.proxy = this.createProxyConfigFromUrl(proxyUrl);
      }

      const res = await this.req(options);

      const end = process.hrtime.bigint();
      timings.api.push(Number(end - start));

      return options.responseType === 'stream' || options.fullResponse ?
        res as any :
        res.data;
    }
    catch (error: any) {
      const end = process.hrtime.bigint();
      timings.api.push(Number(end - start));

      if (error && error.response && (error.response.status === 429 || error.response.status === 503)) {
        let retryAfter: number = parseInt(error.response.headers['retry-after'] || '10');

        if (isNaN(retryAfter)) {
          retryAfter = 10;
        }

        if (this._debug) {
          await (this._logger as Logger).log(`Request throttled. Waiting ${retryAfter} sec before retrying...`);
        }

        await timersUtil.setTimeout(retryAfter * 1000);
        return this.execute(options);
      }

      throw error;
    }
  }

  private updateRequestForCloudType(options: AxiosRequestConfig, cloudType: CloudType): void {
    const url = new URL(options.url!);
    const hostname = `${url.protocol}//${url.hostname}`;
    const cloudUrl: string = Auth.getEndpointForResource(hostname, cloudType);
    options.url = options.url!.replace(hostname, cloudUrl);
  }

  private removeDoubleSlashes(options: AxiosRequestConfig): void {
    options.url = options.url!.substring(0, 8) +
      options.url!.substring(8).replace('//', '/');
  }

  private createProxyConfigFromUrl(url: string): AxiosProxyConfig {
    const parsedUrl = new URL(url);
    const port = parsedUrl.port || (url.toLowerCase().startsWith('https') ? 443 : 80);
    let authObject = null;
    if (parsedUrl.username && parsedUrl.password) {
      authObject = {
        username: parsedUrl.username,
        password: parsedUrl.password
      };
    }
    return { host: parsedUrl.hostname, port: Number(port), protocol: 'http', ...(authObject && { auth: authObject }) };
  }
}

export default new Request();
