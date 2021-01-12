import type { AuthenticationContext, ErrorResponse, Logging, LoggingLevel, TokenResponse, UserCodeInfo } from 'adal-node';
import type * as NodeForge from 'node-forge';
import { FileTokenStorage } from './auth/FileTokenStorage';
import { TokenStorage } from './auth/TokenStorage';
import { Logger } from './cli';
import { CommandError } from './Command';
import config from './config';
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
  certificateType: CertificateType = CertificateType.Unknown;
  certificate?: string
  thumbprint?: string;
  accessTokens: Hash<AccessToken>;
  spoUrl?: string;
  tenantId?: string;
  // ID of the Azure AD app used to authenticate
  appId: string;
  // ID of the tenant where the Azure AD app is registered; common if multitenant
  tenant: string;

  constructor() {
    this.accessTokens = {};
    this.appId = config.cliAadAppId;
    this.tenant = config.tenant;
  }

  public logout(): void {
    this.connected = false;
    this.accessTokens = {};
    this.refreshToken = undefined;
    this.authType = AuthType.DeviceCode;
    this.userName = undefined;
    this.password = undefined;
    this.certificateType = CertificateType.Unknown;
    this.certificate = undefined;
    this.thumbprint = undefined;
    this.spoUrl = undefined;
    this.tenantId = undefined;
    this.appId = config.cliAadAppId;
    this.tenant = config.tenant;
  }
}

export enum AuthType {
  DeviceCode,
  Password,
  Certificate,
  Identity
}

export enum CertificateType {
  Unknown,
  Base64,
  Binary
}

export class Auth {
  private _authCtx: AuthenticationContext | undefined;
  protected get authCtx(): AuthenticationContext {
    if (!this._authCtx) {
      const authenticationContext: typeof AuthenticationContext = require('adal-node').AuthenticationContext;
      this._authCtx = new authenticationContext(`https://login.microsoftonline.com/${this.service.tenant}`);
    }

    return this._authCtx;
  };
  private userCodeInfo?: UserCodeInfo;
  private _service: Service;

  public get service(): Service {
    return this._service;
  }

  public get defaultResource(): string {
    return 'https://graph.microsoft.com';
  }

  constructor() {
    this._service = new Service();
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

  public ensureAccessToken(resource: string, logger: Logger, debug: boolean = false, fetchNew: boolean = false): Promise<string> {
    const logging: typeof Logging = require('adal-node').Logging;
    logging.setLoggingOptions({
      level: debug ? 3 : 0,
      log:
        // Dependency code, we do not control when and how it gets called, thus ignored from coverage
        /* c8 ignore next 3 */
        (level: LoggingLevel, message: string, error?: Error): void => {
          logger.log(message);
        }
    });

    return new Promise<string>((resolve: (accessToken: string) => void, reject: (error: any) => void): void => {
      const now: Date = new Date();
      const accessToken: AccessToken | undefined = this.service.accessTokens[resource];
      const expiresOn: Date = accessToken ? new Date(accessToken.expiresOn) : new Date(0);

      if (!fetchNew && accessToken && expiresOn > now) {
        if (debug) {
          logger.logToStderr(`Existing access token ${accessToken.value} still valid. Returning...`);
        }
        resolve(accessToken.value);
        return;
      }
      else {
        if (debug) {
          if (!accessToken) {
            logger.logToStderr(`No token found for resource ${resource}`);
          }
          else {
            logger.logToStderr(`Access token expired. Token: ${accessToken.value}, ExpiresAt: ${accessToken.expiresOn}`);
          }
        }
      }

      let getTokenPromise: (resource: string, logger: Logger, debug: boolean) => Promise<TokenResponse> = this.ensureAccessTokenWithDeviceCode.bind(this);

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

      getTokenPromise(resource, logger, debug)
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
            logger.logToStderr(new CommandError(_error));
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

  private ensureAccessTokenWithRefreshToken(resource: string, logger: Logger, debug: boolean): Promise<TokenResponse> {
    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (error: any) => void): void => {
      if (debug) {
        logger.logToStderr(`Retrieving new access token using existing refresh token ${this.service.refreshToken}`);
      }

      this.authCtx.acquireTokenWithRefreshToken(
        this.service.refreshToken as string,
        this.service.appId,
        resource,
        (error: Error, response: TokenResponse | ErrorResponse): void => {
          if (debug) {
            logger.logToStderr('Response:');
            logger.logToStderr(response);
            logger.logToStderr('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          resolve(<TokenResponse>response);
        });
    });
  }

  private ensureAccessTokenWithDeviceCode(resource: string, logger: Logger, debug: boolean): Promise<TokenResponse> {
    if (debug) {
      logger.logToStderr(`Starting Auth.ensureAccessTokenWithDeviceCode. resource: ${resource}, debug: ${debug}`);
    }

    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (err: any) => void) => {
      if (debug) {
        logger.logToStderr('No existing refresh token. Starting new device code flow...');
      }

      this.authCtx.acquireUserCode(resource, this.service.appId, 'en-us',
        (error: Error, response: UserCodeInfo): void => {
          if (debug) {
            logger.logToStderr('Response:');
            logger.logToStderr(response);
            logger.logToStderr('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          logger.log(response.message);

          this.userCodeInfo = response;
          this.authCtx.acquireTokenWithDeviceCode(resource, this.service.appId, response,
            (error: Error, response: TokenResponse | ErrorResponse): void => {
              if (debug) {
                logger.logToStderr('Response:');
                logger.logToStderr(response);
                logger.logToStderr('');
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

  private ensureAccessTokenWithPassword(resource: string, logger: Logger, debug: boolean): Promise<TokenResponse> {
    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (error: any) => void): void => {
      if (debug) {
        logger.logToStderr(`Retrieving new access token using credentials...`);
      }

      this.authCtx.acquireTokenWithUsernamePassword(
        resource,
        this.service.userName as string,
        this.service.password as string,
        this.service.appId,
        (error: Error, response: TokenResponse | ErrorResponse): void => {
          if (debug) {
            logger.logToStderr('Response:');
            logger.logToStderr(response);
            logger.logToStderr('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          resolve(<TokenResponse>response);
        });
    });
  }

  private ensureAccessTokenWithCertificate(resource: string, logger: Logger, debug: boolean): Promise<TokenResponse> {
    const nodeForge: typeof NodeForge = require('node-forge');
    const { pem, pki, asn1, pkcs12 } = nodeForge;
    
    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (error: any) => void): void => {
      if (debug) {
        logger.logToStderr(`Retrieving new access token using certificate...`);
      }

      let cert: string = '';
      const buf = Buffer.from(this.service.certificate as string, 'base64');

      if (this.service.certificateType === CertificateType.Unknown || this.service.certificateType === CertificateType.Base64) {
        // First time this method is called, we don't know if certificate is PEM or PFX (type is Unknown)
        // We assume it is PEM but when parsing of PEM fails, we assume it could be PFX
        // Type is persisted on service so subsequent calls only run through the correct parsing flow
        try {
          cert = buf.toString('utf8');
          const pemObjs: NodeForge.pem.ObjectPEM[] = pem.decode(cert);

          if (this.service.thumbprint === undefined) {
            const pemCertObj: NodeForge.pem.ObjectPEM = pemObjs.find(pem => pem.type === "CERTIFICATE") as NodeForge.pem.ObjectPEM;
            const pemCertStr: string = pem.encode(pemCertObj);
            const pemCert: NodeForge.pki.Certificate = pki.certificateFromPem(pemCertStr);

            this.service.thumbprint = this.calculateThumbprint(pemCert);
          }
        }
        catch (e) {
          this.service.certificateType = CertificateType.Binary;
        }
      }

      if (this.service.certificateType === CertificateType.Binary) {
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

          logger.logToStderr(`pkcs8ShroudedKeyBagkeyBags length is ${[pki.oids.pkcs8ShroudedKeyBag].length}`);

          keyBags = p12Parsed.getBags({ bagType: pki.oids.keyBag });
          logger.logToStderr(`keyBag length is ${keyBags[pki.oids.keyBag].length}`);
        }

        // convert a Forge private key to an ASN.1 RSAPrivateKey
        const rsaPrivateKey = pki.privateKeyToAsn1(pkcs8ShroudedKeyBag.key);

        // wrap an RSAPrivateKey ASN.1 object in a PKCS#8 ASN.1 PrivateKeyInfo
        const privateKeyInfo = pki.wrapRsaPrivateKey(rsaPrivateKey);

        // convert a PKCS#8 ASN.1 PrivateKeyInfo to PEM
        cert = pki.privateKeyInfoToPem(privateKeyInfo);

        if (this.service.thumbprint === undefined) {
          const certBags: {
            [key: string]: NodeForge.pkcs12.Bag[] | undefined;
            localKeyId?: NodeForge.pkcs12.Bag[];
            friendlyName?: NodeForge.pkcs12.Bag[];
          } = p12Parsed.getBags({ bagType: pki.oids.certBag });
          const certBag: NodeForge.pkcs12.Bag = (certBags[pki.oids.certBag] as NodeForge.pkcs12.Bag[])[0];

          this.service.thumbprint = this.calculateThumbprint(certBag.cert as NodeForge.pki.Certificate);
        }
      }

      this.authCtx.acquireTokenWithClientCertificate(
        resource,
        this.service.appId,
        cert as string,
        this.service.thumbprint as string,
        (error: Error, response: TokenResponse | ErrorResponse): void => {
          if (debug) {
            logger.logToStderr('Response:');
            logger.logToStderr(response);
            logger.logToStderr('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          resolve(<TokenResponse>response);
        });
    });
  }

  private ensureAccessTokenWithIdentity(resource: string, logger: Logger, debug: boolean): Promise<TokenResponse> {
    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (error: any) => void): void => {
      const userName = this.service.userName;
      if (debug) {
        logger.logToStderr('Wil try to retrieve access token using identity...');
      }

      const requestOptions: any = {
        url: '',
        headers: {
          accept: 'application/json',
          Metadata: true,
          'x-anonymous': true
        },
        responseType: 'json'
      };

      if (process.env.IDENTITY_ENDPOINT && process.env.IDENTITY_HEADER) {
        if (debug) {
          logger.logToStderr('IDENTITY_ENDPOINT and IDENTITY_HEADER env variables found it is Azure Function, WebApp...');
        }

        requestOptions.url = `${process.env.IDENTITY_ENDPOINT}?resource=${encodeURIComponent(resource)}&api-version=2019-08-01`;
        requestOptions.headers['X-IDENTITY-HEADER'] = process.env.IDENTITY_HEADER;
      }
      else if (process.env.MSI_ENDPOINT && process.env.MSI_SECRET) {
        if (debug) {
          logger.logToStderr('MSI_ENDPOINT and MSI_SECRET env variables found it is Azure Function or WebApp, but using the old names of the env variables...');
        }

        requestOptions.url = `${process.env.MSI_ENDPOINT}?resource=${encodeURIComponent(resource)}&api-version=2019-08-01`;
        requestOptions.headers['X-IDENTITY-HEADER'] = process.env.MSI_SECRET;
      }
      else if (process.env.IDENTITY_ENDPOINT) {
        if (debug) {
          logger.logToStderr('IDENTITY_ENDPOINT env variable found it is Azure Could Shell...');
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
          logger.logToStderr('MSI_ENDPOINT env variable found it is Azure Could Shell, but using the old names of the env variables...');
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
          logger.logToStderr('IDENTITY_ENDPOINT and MSI_ENDPOINT env variables not found. Attempt to get Managed Identity token by using the Azure Virtual Machine API...');
        }

        requestOptions.url = `http://169.254.169.254/metadata/identity/oauth2/token?resource=${encodeURIComponent(resource)}&api-version=2018-02-01`;
      }

      if (userName) {
        // if name present then the identity is user-assigned managed identity
        // the name option in this case is either client_id or principal_id (object_id) 
        // of the managed identity service principal
        requestOptions.url += `&client_id=${encodeURIComponent(userName as string)}`;

        if (debug) {
          logger.logToStderr('Wil try to get token using client_id param...');
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
            logger.logToStderr('Wil try to get token using principal_id (also known as object_id) param ...');
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

  private calculateThumbprint(certificate: NodeForge.pki.Certificate): string {
    const nodeForge: typeof NodeForge = require('node-forge');
    const { md, asn1, pki } = nodeForge;

    const messageDigest: NodeForge.md.MessageDigest = md.sha1.create();
    messageDigest.update(asn1.toDer(pki.certificateToAsn1(certificate)).getBytes());
    return messageDigest.digest().toHex();
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

  public static isAppOnlyAuth(accessToken: string): boolean | undefined {
    let isAppOnlyAuth: boolean | undefined;

    if (!accessToken || accessToken.length === 0) {
      return isAppOnlyAuth;
    }

    const chunks = accessToken.split('.');
    if (chunks.length !== 3) {
      return isAppOnlyAuth;
    }

    const tokenString: string = Buffer.from(chunks[1], 'base64').toString();
    try {
      const token: any = JSON.parse(tokenString);
      isAppOnlyAuth = !token.upn;
    }
    catch {
    }

    return isAppOnlyAuth;
  }
}

export default new Auth();