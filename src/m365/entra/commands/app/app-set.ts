import { Application, KeyCredential, PublicClientApplication, SpaApplication, WebApplication } from '@microsoft/microsoft-graph-types';
import fs from 'fs';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { zod } from '../../../../utils/zod.js';

const entraIDApplicationPlatform = ['spa', 'web', 'publicClient'] as const;

const options = globalOptionsZod
  .extend({
    appId: z.string().uuid().optional(),
    objectId: z.string().uuid().optional(),
    name: z.string().optional(),
    platform: z.enum(entraIDApplicationPlatform).optional(),
    redirectUris: z.string().optional(),
    redirectUrisToRemove: z.string().optional(),
    uris: z.string().optional(),
    certificateFile: z.string().optional(),
    certificateBase64Encoded: z.string().optional(),
    certificateDisplayName: z.string().optional(),
    allowPublicClientFlows: z.boolean().optional()
  })
  .passthrough();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAppSetCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_SET;
  }

  public get description(): string {
    return 'Updates Entra app registration';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  constructor() {
    super();
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.appId, options.objectId, options.name].filter(Boolean).length === 1, {
        message: 'Specify either appId, objectId, or name but not multiple'
      })
      .refine(options => !options.redirectUris || !!options.platform, {
        message: 'When you specify redirectUris you also need to specify platform'
      })
      .refine(options => !(options.certificateFile && options.certificateBase64Encoded), {
        message: 'Specify either certificateFile or certificateBase64Encoded but not both'
      })
      .refine(options => {
        if (options.certificateDisplayName && !options.certificateFile && !options.certificateBase64Encoded) {
          return false;
        }
        return true;
      }, {
        message: 'When you specify certificateDisplayName you also need to specify certificateFile or certificateBase64Encoded'
      })
      .refine(options => {
        if (options.certificateFile && !fs.existsSync(options.certificateFile)) {
          return false;
        }
        return true;
      }, {
        message: 'Certificate file not found'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let objectId = await this.getAppObjectId(args, logger);
      objectId = await this.updateUnknownOptions(args, objectId);
      objectId = await this.configureUri(args, objectId, logger);
      objectId = await this.configureRedirectUris(args, objectId, logger);
      objectId = await this.updateAllowPublicClientFlows(args, objectId, logger);
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
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : name}...`);
    }

    if (appId) {
      const app = await entraApp.getAppRegistrationByAppId(appId, ['id']);
      return app.id!;
    }
    else {
      const app = await entraApp.getAppRegistrationByAppName(name!, ["id"]);
      return app.id!;
    }
  }

  private async updateUnknownOptions(args: CommandArgs, objectId: string): Promise<string> {
    const unknownOptions = optionsUtils.getUnknownOptions(args.options, zod.schemaToOptions(this.schema!));

    if (Object.keys(unknownOptions).length > 0) {
      const requestBody = {};
      optionsUtils.addUnknownOptionsToPayload(requestBody, unknownOptions);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
        headers: {
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: requestBody
      };
      await request.patch(requestOptions);
    }
    return objectId;
  }

  private async updateAllowPublicClientFlows(args: CommandArgs, objectId: string, logger: Logger): Promise<string> {
    if (args.options.allowPublicClientFlows === undefined) {
      return objectId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Configuring Entra application AllowPublicClientFlows option...`);
    }

    const applicationInfo: any = {
      isFallbackPublicClient: args.options.allowPublicClientFlows
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

  private async configureUri(args: CommandArgs, objectId: string, logger: Logger): Promise<string> {
    if (!args.options.uris) {
      return objectId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Configuring Microsoft Entra application ID URI...`);
    }

    const identifierUris: string[] = args.options.uris
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
      await logger.logToStderr(`Configuring Microsoft Entra application redirect URIs...`);
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
      await logger.logToStderr(`Setting certificate for Microsoft Entra app...`);
    }

    const certificateBase64Encoded = await this.getCertificateBase64Encoded(args, logger);

    const currentKeyCredentials = await this.getCurrentKeyCredentialsList(args, objectId, certificateBase64Encoded, logger);
    if (this.verbose) {
      await logger.logToStderr(`Adding new keyCredential to list`);
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

  private async getCertificateBase64Encoded(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.certificateBase64Encoded) {
      return args.options.certificateBase64Encoded;
    }

    if (this.debug) {
      await logger.logToStderr(`Reading existing ${args.options.certificateFile}...`);
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
      await logger.logToStderr(`Retrieving current keyCredentials list for app`);
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
      await logger.logToStderr(`Updating keyCredentials in Microsoft Entra app`);
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

export default new EntraAppSetCommand();