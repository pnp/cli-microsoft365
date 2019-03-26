import * as os from 'os';
import { TokenStorage } from './auth/TokenStorage';
import { KeychainTokenStorage } from './auth/KeychainTokenStorage';
import { WindowsTokenStorage } from './auth/WindowsTokenStorage';
import { FileTokenStorage } from './auth/FileTokenStorage';
import { AuthenticationContext, TokenResponse, ErrorResponse, UserCodeInfo, Logging, LoggingLevel } from 'adal-node';
import { CommandError } from './Command';
import config from './config';

export interface Hash<TValue> {
  [key: string]: TValue;
}

export interface AccessToken {
  expiresOn: string;
  value: string;
}

export class Service {
  connected: boolean = false;
  refreshToken?: string;
  authType: AuthType = AuthType.DeviceCode;
  userName?: string;
  password?: string;
  certificate?: string
  thumbprint?: string;
  accessTokens: Hash<AccessToken>;
  spoUrl?: string;
  tenantId?: string;

  constructor() {
    this.accessTokens = {};
  }

  public logout(): void {
    this.connected = false;
    this.accessTokens = {};
    this.refreshToken = undefined;
    this.authType = AuthType.DeviceCode;
    this.userName = undefined;
    this.password = undefined;
    this.certificate = undefined;
    this.thumbprint = undefined;
    this.spoUrl = undefined;
    this.tenantId = undefined;
  }
}

export interface Logger {
  log: (msg: any) => void
}

export enum AuthType {
  DeviceCode,
  Password,
  Certificate
}

export class Auth {
  protected authCtx: AuthenticationContext;
  private userCodeInfo?: UserCodeInfo;
  private _service: Service;
  private appId: string;

  public get service(): Service {
    return this._service;
  }

  public get defaultResource(): string {
    return 'https://graph.microsoft.com';
  }

  constructor() {
    this.appId = config.cliAadAppId;
    this._service = new Service();
    this.authCtx = new AuthenticationContext(`https://login.microsoftonline.com/${config.tenant}`);
  }

  public restoreAuth(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .getServiceConnectionInfo<Service>()
        .then((service: Service): void => {
          this._service = Object.assign(this._service, service);
          resolve();
        }, (error: any): void => {
          resolve();
        });
    });
  }

  public ensureAccessToken(resource: string, stdout: Logger, debug: boolean = false): Promise<string> {
    /* istanbul ignore next */
    Logging.setLoggingOptions({
      level: debug ? 3 : 0,
      log: (level: LoggingLevel, message: string, error?: Error): void => {
        stdout.log(message);
      }
    });

    return new Promise<string>((resolve: (accessToken: string) => void, reject: (error: any) => void): void => {
      const now: Date = new Date();
      const accessToken: AccessToken | undefined = this.service.accessTokens[resource];
      const expiresOn: Date = accessToken ? new Date(accessToken.expiresOn) : new Date(0);

      if (accessToken && expiresOn > now) {
        if (debug) {
          stdout.log(`Existing access token ${accessToken.value} still valid. Returning...`);
        }
        resolve(accessToken.value);
        return;
      }
      else {
        if (debug) {
          if (!accessToken) {
            stdout.log(`No token found for resource ${resource}`);
          }
          else {
            stdout.log(`Access token expired. Token: ${accessToken.value}, ExpiresAt: ${accessToken.expiresOn}`);
          }
        }
      }

      let getTokenPromise: (resource: string, stdout: Logger, debug: boolean) => Promise<TokenResponse> = this.ensureAccessTokenWithDeviceCode.bind(this);

      if (this.service.refreshToken) {
        getTokenPromise = this.ensureAccessTokenWithRefreshToken.bind(this);
      }
      else {
        switch (this.service.authType) {
          case AuthType.DeviceCode:
            getTokenPromise = this.ensureAccessTokenWithDeviceCode.bind(this);
            break;
          case AuthType.Password:
            getTokenPromise = this.ensureAccessTokenWithPassword.bind(this);
            break;
          case AuthType.Certificate:
            getTokenPromise = this.ensureAccessTokenWithCertificate.bind(this);
            break;
        }
      }

      let error: any = undefined;

      getTokenPromise(resource, stdout, debug)
        .then((tokenResponse: TokenResponse): Promise<void> => {
          this.service.accessTokens[resource] = {
            expiresOn: tokenResponse.expiresOn as string,
            value: tokenResponse.accessToken
          };
          this.service.refreshToken = tokenResponse.refreshToken;
          this.service.connected = true;
          return this.storeConnectionInfo();
        }, (_error: any): Promise<void> => {
          error = _error;
          // return rejected promise to prevent resolving the parent promise
          // with access token when there is none
          return Promise.reject(_error);
        })
        .then((): void => {
          resolve(this.service.accessTokens[resource].value);
        }, (_error: any): void => {
          // _error could happen due to an issue with persisting the access
          // token which shouldn't fail the overall token retrieval process
          if (debug) {
            stdout.log(new CommandError(_error));
          }
          // was there an issue earlier in the process
          if (error) {
            reject(error);
          }
          else {
            // resolve with the retrieved token despite the issue with
            // persisting the token
            resolve(this.service.accessTokens[resource].value);
          }
        });
    });
  }

  private ensureAccessTokenWithRefreshToken(resource: string, stdout: Logger, debug: boolean): Promise<TokenResponse> {
    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (error: any) => void): void => {
      if (debug) {
        stdout.log(`Retrieving new access token using existing refresh token ${this.service.refreshToken}`);
      }

      this.authCtx.acquireTokenWithRefreshToken(
        this.service.refreshToken as string,
        this.appId as string,
        resource,
        (error: Error, response: TokenResponse | ErrorResponse): void => {
          if (debug) {
            stdout.log('Response:');
            stdout.log(response);
            stdout.log('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          resolve(<TokenResponse>response);
        });
    });
  }

  private ensureAccessTokenWithDeviceCode(resource: string, stdout: Logger, debug: boolean): Promise<TokenResponse> {
    if (debug) {
      stdout.log(`Starting Auth.ensureAccessTokenWithDeviceCode. resource: ${resource}, debug: ${debug}`);
    }

    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (err: any) => void) => {
      if (debug) {
        stdout.log('No existing refresh token. Starting new device code flow...');
      }

      this.authCtx.acquireUserCode(resource, this.appId as string, 'en-us',
        (error: Error, response: UserCodeInfo): void => {
          if (debug) {
            stdout.log('Response:');
            stdout.log(response);
            stdout.log('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          stdout.log(response.message);

          this.userCodeInfo = response;
          this.authCtx.acquireTokenWithDeviceCode(resource, this.appId as string, response,
            (error: Error, response: TokenResponse | ErrorResponse): void => {
              if (debug) {
                stdout.log('Response:');
                stdout.log(response);
                stdout.log('');
              }

              if (error) {
                reject((response && (response as any).error_description) || error.message);
                return;
              }

              this.userCodeInfo = undefined;
              resolve(<TokenResponse>response);
            });
        });
    });
  }

  private ensureAccessTokenWithPassword(resource: string, stdout: Logger, debug: boolean): Promise<TokenResponse> {
    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (error: any) => void): void => {
      if (debug) {
        stdout.log(`Retrieving new access token using credentials...`);
      }

      this.authCtx.acquireTokenWithUsernamePassword(
        resource,
        this.service.userName as string,
        this.service.password as string,
        this.appId as string,
        (error: Error, response: TokenResponse | ErrorResponse): void => {
          if (debug) {
            stdout.log('Response:');
            stdout.log(response);
            stdout.log('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          resolve(<TokenResponse>response);
        });
    });
  }

  private ensureAccessTokenWithCertificate(resource: string, stdout: Logger, debug: boolean): Promise<TokenResponse> {
    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (error: any) => void): void => {
      if (debug) {
        stdout.log(`Retrieving new access token using certificate (thumbprint ${this.service.thumbprint})...`);
      }

      this.authCtx.acquireTokenWithClientCertificate(
        resource,
        this.appId as string,
        this.service.certificate as string,
        this.service.thumbprint as string,
        (error: Error, response: TokenResponse | ErrorResponse): void => {
          if (debug) {
            stdout.log('Response:');
            stdout.log(response);
            stdout.log('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          resolve(<TokenResponse>response);
        });
    });
  }

  public cancel(): void {
    if (this.userCodeInfo) {
      this.authCtx.cancelRequestToGetTokenWithDeviceCode(this.userCodeInfo as UserCodeInfo, /* istanbul ignore next */(error: Error, response: TokenResponse | ErrorResponse): void => { });
    }
  }

  public static getResourceFromUrl(url: string): string {
    let resource: string = url;
    const pos: number = resource.indexOf('/', 8);
    if (pos > -1) {
      resource = resource.substr(0, pos);
    }

    return resource;
  }

  private getServiceConnectionInfo<TConn>(): Promise<TConn> {
    return new Promise<TConn>((resolve: (connectionInfo: TConn) => void, reject: (error: any) => void): void => {
      const tokenStorage = this.getTokenStorage();
      tokenStorage
        .get()
        .then((json: string): void => {
          try {
            resolve(JSON.parse(json));
          }
          catch (err) {
            reject(err);
          }
        }, (error: any): void => {
          reject(error);
        })
    });
  }

  public storeConnectionInfo(): Promise<void> {
    const tokenStorage = this.getTokenStorage();
    return tokenStorage.set(JSON.stringify(this.service));
  }

  public clearConnectionInfo(): Promise<void> {
    const tokenStorage = this.getTokenStorage();
    return tokenStorage.remove();
  }

  public getTokenStorage(): TokenStorage {
    const platform: NodeJS.Platform = os.platform();
    let tokenStorage: TokenStorage;
    switch (platform) {
      case 'darwin':
        tokenStorage = new KeychainTokenStorage();
        break;
      case 'win32':
        tokenStorage = new WindowsTokenStorage();
        break;
      default:
        tokenStorage = new FileTokenStorage();
        break;
    }

    return tokenStorage;
  }
}

export default new Auth();