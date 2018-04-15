import * as os from 'os';
import { TokenStorage } from './auth/TokenStorage';
import { KeychainTokenStorage } from './auth/KeychainTokenStorage';
import { WindowsTokenStorage } from './auth/WindowsTokenStorage';
import { FileTokenStorage } from './auth/FileTokenStorage';
import { AuthenticationContext, TokenResponse, ErrorResponse, UserCodeInfo, Logging, LoggingLevel } from 'adal-node';
import { CommandError } from './Command';

export class Service {
  connected: boolean = false;
  resource: string;
  accessToken: string = '';
  refreshToken?: string;
  expiresOn: string = '';
  authType: AuthType = AuthType.DeviceCode;
  userName?: string;
  password?: string;

  constructor(resource: string = '') {
    this.resource = resource;
  }

  public disconnect(): void {
    this.connected = false;
    this.resource = '';
    this.accessToken = '';
    this.refreshToken = undefined;
    this.expiresOn = '';
    this.authType = AuthType.DeviceCode;
    this.userName = undefined;
    this.password = undefined;
  }
}

export interface Logger {
  log: (msg: any) => void
}

export enum AuthType {
  DeviceCode,
  Password
}

export abstract class Auth {
  protected authCtx: AuthenticationContext;
  private userCodeInfo?: UserCodeInfo;

  protected abstract serviceId(): string;

  constructor(public service: Service, private appId?: string) {
    this.authCtx = new AuthenticationContext('https://login.microsoftonline.com/common');
  }

  public restoreAuth(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .getServiceConnectionInfo<Service>(this.serviceId())
        .then((service: Service): void => {
          this.service = service;
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
      const expiresOn: Date = new Date(this.service.expiresOn);

      if (expiresOn > now &&
        this.service.accessToken !== undefined) {
        if (debug) {
          stdout.log(`Existing access token ${this.service.accessToken} still valid. Returning...`);
        }
        resolve(this.service.accessToken);
        return;
      }
      else {
        if (debug) {
          stdout.log(`No existing access token or expired. Token: ${this.service.accessToken}, ExpiresAt: ${this.service.expiresOn}`);
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
        }
      }

      getTokenPromise(resource, stdout, debug)
        .then((tokenResponse: TokenResponse): Promise<void> => {
          this.service.accessToken = tokenResponse.accessToken;
          this.service.refreshToken = tokenResponse.refreshToken;
          this.service.expiresOn = tokenResponse.expiresOn as string;
          return this.setServiceConnectionInfo(this.serviceId(), this.service);
        }, (error: any): void => {
          reject(error);
        })
        .then((): void => {
          resolve(this.service.accessToken);
        }, (error: any): void => {
          if (debug) {
            stdout.log(new CommandError(error));
          }
          resolve(this.service.accessToken);
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

  public cancel(): void {
    if (this.userCodeInfo) {
      this.authCtx.cancelRequestToGetTokenWithDeviceCode(this.userCodeInfo as UserCodeInfo, /* istanbul ignore next */(error: Error, response: TokenResponse | ErrorResponse): void => { });
    }
  }

  public getAccessToken(resource: string, refreshToken: string, stdout: Logger, debug: boolean = false): Promise<string> {
    if (debug) {
      stdout.log(`Starting Auth.getAccessToken. resource: ${resource}, refreshToken: ${refreshToken}, debug: ${debug}`);
    }

    return new Promise<string>((resolve: (accessToken: string) => void, reject: (err: any) => void): void => {
      if (debug) {
        stdout.log(`Retrieving access token for ${resource} using refresh token ${refreshToken}`);
      }

      this.authCtx.acquireTokenWithRefreshToken(refreshToken, this.appId as string, resource,
        (error: Error, response: TokenResponse | ErrorResponse): void => {
          if (error) {
            if (debug) {
              stdout.log('Error:');
              stdout.log(error);
              stdout.log('');
            }

            reject((response && (response as any).error_description) || error.message);
            return;
          }

          if (debug) {
            stdout.log('Response:');
            stdout.log(response);
            stdout.log('');
          }

          const tokenResponse: TokenResponse = <TokenResponse>response;
          resolve(tokenResponse.accessToken);
        });
    });
  }

  public static getResourceFromUrl(url: string): string {
    let resource: string = url;
    const pos: number = resource.indexOf('/', 8);
    if (pos > -1) {
      resource = resource.substr(0, pos);
    }

    return resource;
  }

  protected getServiceConnectionInfo<TConn>(service: string): Promise<TConn> {
    return new Promise<TConn>((resolve: (connectionInfo: TConn) => void, reject: (error: any) => void): void => {
      const tokenStorage = this.getTokenStorage();
      tokenStorage
        .get(service)
        .then((json: string): void => {
          resolve(JSON.parse(json));
        }, (error: any): void => {
          reject(error);
        })
    });
  }

  protected setServiceConnectionInfo<TConn>(service: string, connectionInfo: TConn): Promise<void> {
    const tokenStorage = this.getTokenStorage();
    return tokenStorage.set(service, JSON.stringify(connectionInfo));
  }

  public storeConnectionInfo(): Promise<void> {
    return this.setServiceConnectionInfo(this.serviceId(), this.service);
  }

  protected clearServiceConnectionInfo(service: string): Promise<void> {
    const tokenStorage = this.getTokenStorage();
    return tokenStorage.remove(service);
  }

  public clearConnectionInfo(): Promise<void> {
    return this.clearServiceConnectionInfo(this.serviceId());
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