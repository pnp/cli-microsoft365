import { TokenStorage } from './auth/TokenStorage';
import { FileTokenStorage } from './auth/FileTokenStorage';
import { AuthenticationContext, TokenResponse, ErrorResponse, UserCodeInfo, Logging, LoggingLevel } from 'adal-node';
import { CommandError } from './Command';
import config from './config';
import { asn1, pkcs12, pki } from 'node-forge';
import request from './request';

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
  Certificate,
  Identity
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

  public ensureAccessToken(resource: string, stdout: Logger, debug: boolean = false, fetchNew: boolean = false): Promise<string> {
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

      if (!fetchNew && accessToken && expiresOn > now) {
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
          case AuthType.Identity:
            getTokenPromise = this.ensureAccessTokenWithIdentity.bind(this);
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
                reject((response && (response as any).error_description) || error.message || (error as any).error_description);
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

      let cert: string = '';

      if (this.service.password === undefined) {
        const buf = Buffer.from(this.service.certificate as string, 'base64');
        cert = buf.toString('utf8');
      }
      else {
        const buf = Buffer.from(this.service.certificate as string, 'base64');
        const p12Asn1 = asn1.fromDer(buf.toString('binary'), false);

        const p12Parsed = pkcs12.pkcs12FromAsn1(p12Asn1, false, this.service.password);

        let keyBags: any = p12Parsed.getBags({ bagType: pki.oids.pkcs8ShroudedKeyBag });
        const pkcs8ShroudedKeyBag = keyBags[pki.oids.pkcs8ShroudedKeyBag][0];

        if (debug) {
          // check if there is something in the keyBag as well as
          // the pkcs8ShroudedKeyBag. This will give us more information
          // whether there is a cert that can potentially store keys in the keyBag.
          // I could not find a way to add something to the keyBag with all 
          // my attempts, but lets keep it here for troubleshooting purposes.

          stdout.log(`pkcs8ShroudedKeyBagkeyBags length is ${[pki.oids.pkcs8ShroudedKeyBag].length}`);

          keyBags = p12Parsed.getBags({ bagType: pki.oids.keyBag });
          stdout.log(`keyBag length is ${keyBags[pki.oids.keyBag].length}`);
        }

        // convert a Forge private key to an ASN.1 RSAPrivateKey
        const rsaPrivateKey = pki.privateKeyToAsn1(pkcs8ShroudedKeyBag.key);

        // wrap an RSAPrivateKey ASN.1 object in a PKCS#8 ASN.1 PrivateKeyInfo
        const privateKeyInfo = pki.wrapRsaPrivateKey(rsaPrivateKey);

        // convert a PKCS#8 ASN.1 PrivateKeyInfo to PEM
        cert = pki.privateKeyInfoToPem(privateKeyInfo);
      }

      this.authCtx.acquireTokenWithClientCertificate(
        resource,
        this.appId as string,
        cert as string,
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

  private ensureAccessTokenWithIdentity(resource: string, stdout: Logger, debug: boolean): Promise<TokenResponse> {
    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (error: any) => void): void => {
      const userName = this.service.userName;
      if (debug) {
        stdout.log('Wil try to retrieve access token using identity...');
      }

      const requestOptions: any = {
        url: '',
        headers: {
          accept: 'application/json',
          Metadata: true,
          'x-anonymous': true
        },
        json: true
      };

      if (process.env.IDENTITY_ENDPOINT && process.env.IDENTITY_HEADER) {
        if (debug) {
          stdout.log('IDENTITY_ENDPOINT and IDENTITY_HEADER env variables found it is Azure Function, WebApp...');
        }

        requestOptions.url = `${process.env.IDENTITY_ENDPOINT}?resource=${encodeURIComponent(resource)}&api-version=2019-08-01`;
        requestOptions.headers['X-IDENTITY-HEADER'] = process.env.IDENTITY_HEADER;
      }
      else if (process.env.MSI_ENDPOINT && process.env.MSI_SECRET) {
        if (debug) {
          stdout.log('MSI_ENDPOINT and MSI_SECRET env variables found it is Azure Function or WebApp, but using the old names of the env variables...');
        }

        requestOptions.url = `${process.env.MSI_ENDPOINT}?resource=${encodeURIComponent(resource)}&api-version=2019-08-01`;
        requestOptions.headers['X-IDENTITY-HEADER'] = process.env.MSI_SECRET;
      }
      else if (process.env.IDENTITY_ENDPOINT) {
        if (debug) {
          stdout.log('IDENTITY_ENDPOINT env variable found it is Azure Could Shell...');
        }

        if (userName && process.env.ACC_CLOUD) {
          // reject for now since the Azure Cloud Shell does not support user-managed identity 
          reject("Azure Cloud Shell does not support user-managed identity. You can execute the command without the --userName option to login with user identity");
          return;
        }

        requestOptions.url = `${process.env.IDENTITY_ENDPOINT}?resource=${encodeURIComponent(resource)}`;

      }
      else if (process.env.MSI_ENDPOINT) {
        if (debug) {
          stdout.log('MSI_ENDPOINT env variable found it is Azure Could Shell, but using the old names of the env variables...');
        }

        if (userName && process.env.ACC_CLOUD) {
          // reject for now since the Azure Cloud Shell does not support user-managed identity 
          reject("Azure Cloud Shell does not support user-managed identity. You can execute the command without the --userName option to login with user identity");
          return;
        }

        requestOptions.url = `${process.env.MSI_ENDPOINT}?resource=${encodeURIComponent(resource)}`;
      }
      else {
        if (debug) {
          stdout.log('IDENTITY_ENDPOINT and MSI_ENDPOINT env variables not found. Attempt to get Managed Identity token by using the Azure Virtual Machine API...');
        }

        requestOptions.url = `http://169.254.169.254/metadata/identity/oauth2/token?resource=${encodeURIComponent(resource)}&api-version=2018-02-01`;
      }

      if (userName) {
        // if name present then the identity is user-assigned managed identity
        // the name option in this case is either client_id or principal_id (object_id) 
        // of the managed identity service principal
        requestOptions.url += `&client_id=${encodeURIComponent(userName as string)}`;

        if (debug) {
          stdout.log('Wil try to get token using client_id param...');
        }
      }

      request
        .get(requestOptions)
        .then((res: any): void => {

          resolve({ accessToken: res.access_token, expiresOn: parseInt(res.expires_on) * 1000 } as any);
          return;
        })
        .catch((e: any) => {
          if (!userName) {
            reject(e);
            return;
          }

          // since the userName option can be either client_id or principal_id (object_id) 
          // and the first attempt was using client_id
          // now lets see if the api returned 'not found' response and
          // try to get token using principal_id (object_id)

          let isNotFoundResponse = false;
          if (e.error && e.error.Message) {
            // check if it is Azure Function api 'not found' response
            isNotFoundResponse = (e.error.Message.indexOf("No Managed Identity found") !== -1);
          }
          else if (e.error && e.error.error_description) {
            // check if it is Azure VM api 'not found' response
            isNotFoundResponse = (e.error.error_description === "Identity not found");
          }

          if (!isNotFoundResponse) {
            // it is not a 'not found' response then exit with error
            reject(e);
            return;
          }

          if (debug) {
            stdout.log('Wil try to get token using principal_id (also known as object_id) param ...');
          }

          requestOptions.url = requestOptions.url.replace('&client_id=', '&principal_id=');
          requestOptions.headers['x-anonymous'] = true;

          request
            .get(requestOptions)
            .then((res: any): void => {
              resolve({ accessToken: res.access_token, expiresOn: parseInt(res.expires_on) * 1000 } as any);
            })
            .catch((err: any) => {
              // will give up and not try any further with the 'msi_res_id' (resource id) query string param
              // since it does not work with the Azure Functions api, but just with the Azure VM api
              if (err.error.code === 'EACCES') {
                // the CLI does not know if managed identity is actually assigned when EACCES code thrown
                // so show meaningful message since the raw error response could be misleading 
                reject('Error while logging with Managed Identity. Please check if a Managed Identity is assigned to the current Azure resource.');
              }
              else {
                reject(err);
              }
            });
        });
    });
  }

  public cancel(): void {
    if (this.userCodeInfo) {
      this.authCtx.cancelRequestToGetTokenWithDeviceCode(this.userCodeInfo as UserCodeInfo, (error: Error, response: TokenResponse | ErrorResponse): void => { });
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
    return new FileTokenStorage();
  }
}

export default new Auth();