import { Application, KeyCredential, PublicClientApplication, SpaApplication, WebApplication } from '@microsoft/microsoft-graph-types';
import * as fs from 'fs';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  objectId?: string;
  name?: string;
  platform?: string;
  redirectUris?: string;
  redirectUrisToRemove?: string;
  uri?: string;
  certificateFile?: string;
  certificateBase64Encoded?: string;
  certificateDisplayName?: string;
}

class AadAppSetCommand extends GraphCommand {
  private static aadApplicationPlatform: string[] = ['spa', 'web', 'publicClient'];

  public get name(): string {
    return commands.APP_SET;
  }

  public get description(): string {
    return 'Updates Azure AD app registration';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        appId: typeof args.options.appId !== 'undefined',
        objectId: typeof args.options.objectId !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        platform: typeof args.options.platform !== 'undefined',
        redirectUris: typeof args.options.redirectUris !== 'undefined',
        redirectUrisToRemove: typeof args.options.redirectUrisToRemove !== 'undefined',
        uri: typeof args.options.uri !== 'undefined',
        certificateFile: typeof args.options.certificateFile !== 'undefined',
        certificateBase64Encoded: typeof args.options.certificateBase64Encoded !== 'undefined',
        certificateDisplayName: typeof args.options.certificateDisplayName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' },
      { option: '--objectId [objectId]' },
      { option: '-n, --name [name]' },
      { option: '-u, --uri [uri]' },
      { option: '-r, --redirectUris [redirectUris]' },
      { option: '--certificateFile [certificateFile]' },
      { option: '--certificateBase64Encoded [certificateBase64Encoded]' },
      { option: '--certificateDisplayName [certificateDisplayName]' },
      {
        option: '--platform [platform]',
        autocomplete: AadAppSetCommand.aadApplicationPlatform
      },
      { option: '--redirectUrisToRemove [redirectUrisToRemove]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.certificateFile && args.options.certificateBase64Encoded) {
          return 'Specify either certificateFile or certificateBase64Encoded but not both';
        }

        if (args.options.certificateDisplayName && !args.options.certificateFile && !args.options.certificateBase64Encoded) {
          return 'When you specify certificateDisplayName you also need to specify certificateFile or certificateBase64Encoded';
        }

        if (args.options.certificateFile && !fs.existsSync(args.options.certificateFile as string)) {
          return 'Certificate file not found';
        }

        if (args.options.redirectUris && !args.options.platform) {
          return `When you specify redirectUris you also need to specify platform`;
        }

        if (args.options.platform &&
          AadAppSetCommand.aadApplicationPlatform.indexOf(args.options.platform) < 0) {
          return `${args.options.platform} is not a valid value for platform. Allowed values are ${AadAppSetCommand.aadApplicationPlatform.join(', ')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'objectId', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let objectId = await this.getAppObjectId(args, logger);
      objectId = await this.configureUri(args, objectId, logger);
      objectId = await this.configureRedirectUris(args, objectId, logger);
      await this.configureCertificate(args, objectId, logger);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.objectId) {
      return args.options.objectId;
    }

    const { appId, name } = args.options;

    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Azure AD app ${appId ? appId : name}...`);
    }

    const filter: string = appId ?
      `appId eq '${formatting.encodeQueryParameter(appId)}'` :
      `displayName eq '${formatting.encodeQueryParameter(name as string)}'`;

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=${filter}&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string }[] }>(requestOptions);

    if (res.value.length === 1) {
      return res.value[0].id;
    }

    if (res.value.length === 0) {
      const applicationIdentifier = appId ? `ID ${appId}` : `name ${name}`;
      throw `No Azure AD application registration with ${applicationIdentifier} found`;
    }

    throw `Multiple Azure AD application registration with name ${name} found. Please disambiguate (app object IDs): ${res.value.map(a => a.id).join(', ')}`;

  }

  private async configureUri(args: CommandArgs, objectId: string, logger: Logger): Promise<string> {
    if (!args.options.uri) {
      return objectId;
    }

    if (this.verbose) {
      logger.logToStderr(`Configuring Azure AD application ID URI...`);
    }

    const identifierUris: string[] = args.options.uri
      .split(',')
      .map(u => u.trim());

    const applicationInfo: any = {
      identifierUris: identifierUris
    };

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: applicationInfo
    };

    await request.patch(requestOptions);
    return objectId;
  }

  private async configureRedirectUris(args: CommandArgs, objectId: string, logger: Logger): Promise<string> {
    if (!args.options.redirectUris && !args.options.redirectUrisToRemove) {
      return objectId;
    }

    if (this.verbose) {
      logger.logToStderr(`Configuring Azure AD application redirect URIs...`);
    }

    const getAppRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const application = await request.get<Application>(getAppRequestOptions);

    const publicClientRedirectUris: string[] = (application.publicClient as PublicClientApplication).redirectUris as string[];
    const spaRedirectUris: string[] = (application.spa as SpaApplication).redirectUris as string[];
    const webRedirectUris: string[] = (application.web as WebApplication).redirectUris as string[];

    // start with existing redirect URIs
    const applicationPatch: Application = {
      publicClient: {
        redirectUris: publicClientRedirectUris
      },
      spa: {
        redirectUris: spaRedirectUris
      },
      web: {
        redirectUris: webRedirectUris
      }
    };

    if (args.options.redirectUrisToRemove) {
      // remove redirect URIs from all platforms
      const redirectUrisToRemove: string[] = args.options.redirectUrisToRemove
        .split(',')
        .map(u => u.trim());

      (applicationPatch.publicClient as PublicClientApplication).redirectUris =
        publicClientRedirectUris.filter(u => !redirectUrisToRemove.includes(u));
      (applicationPatch.spa as SpaApplication).redirectUris =
        spaRedirectUris.filter(u => !redirectUrisToRemove.includes(u));
      (applicationPatch.web as WebApplication).redirectUris =
        webRedirectUris.filter(u => !redirectUrisToRemove.includes(u));
    }

    if (args.options.redirectUris) {
      const urlsToAdd: string[] = args.options.redirectUris
        .split(',')
        .map(u => u.trim());

      // add new redirect URIs. If the URI is already present, it will be ignored
      switch (args.options.platform) {
        case 'spa':
          ((applicationPatch.spa as SpaApplication).redirectUris as string[])
            .push(...urlsToAdd.filter(u => !spaRedirectUris.includes(u)));
          break;
        case 'publicClient':
          ((applicationPatch.publicClient as PublicClientApplication).redirectUris as string[])
            .push(...urlsToAdd.filter(u => !publicClientRedirectUris.includes(u)));
          break;
        case 'web':
          ((applicationPatch.web as WebApplication).redirectUris as string[])
            .push(...urlsToAdd.filter(u => !webRedirectUris.includes(u)));
      }
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: applicationPatch
    };

    await request.patch(requestOptions);
    return objectId;
  }

  private async configureCertificate(args: CommandArgs, objectId: string, logger: Logger): Promise<void> {
    if (!args.options.certificateFile && !args.options.certificateBase64Encoded) {
      return;
    }

    if (this.verbose) {
      logger.logToStderr(`Setting certificate for Azure AD app...`);
    }

    const certificateBase64Encoded = this.getCertificateBase64Encoded(args, logger);

    const currentKeyCredentials = await this.getCurrentKeyCredentialsList(args, objectId, certificateBase64Encoded, logger);
    if (this.verbose) {
      logger.logToStderr(`Adding new keyCredential to list`);
    }

    // The KeyCredential graph type defines the 'key' property as 'NullableOption<number>'
    // while it is a base64 encoded string. This is why a cast to any is used here.
    const keyCredentials = currentKeyCredentials.filter(existingCredential => existingCredential.key !== certificateBase64Encoded as any);

    const newKeyCredential = {
      type: "AsymmetricX509Cert",
      usage: "Verify",
      displayName: args.options.certificateDisplayName,
      key: certificateBase64Encoded
    } as any;

    keyCredentials.push(newKeyCredential);

    return this.updateKeyCredentials(objectId, keyCredentials, logger);
  }

  private getCertificateBase64Encoded(args: CommandArgs, logger: Logger): string {
    if (args.options.certificateBase64Encoded) {
      return args.options.certificateBase64Encoded;
    }

    if (this.debug) {
      logger.logToStderr(`Reading existing ${args.options.certificateFile}...`);
    }

    try {
      return fs.readFileSync(args.options.certificateFile as string, { encoding: 'base64' });
    }
    catch (e) {
      throw new Error(`Error reading certificate file: ${e}. Please add the certificate using base64 option '--certificateBase64Encoded'.`);
    }
  }

  // We first retrieve existing certificates because we need to specify the full list of certificates when updating the app.
  private async getCurrentKeyCredentialsList(args: CommandArgs, objectId: string, certificateBase64Encoded: string, logger: Logger): Promise<KeyCredential[]> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving current keyCredentials list for app`);
    }

    const getAppRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${objectId}?$select=keyCredentials`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const application = await request.get<Application>(getAppRequestOptions);
    return application.keyCredentials || [];
  }

  private async updateKeyCredentials(objectId: string, keyCredentials: KeyCredential[], logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Updating keyCredentials in AAD app`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        keyCredentials: keyCredentials
      }
    };

    return request.patch(requestOptions);
  }
}

module.exports = new AadAppSetCommand();