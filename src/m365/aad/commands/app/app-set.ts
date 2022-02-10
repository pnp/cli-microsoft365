import { Application, PublicClientApplication, SpaApplication, WebApplication } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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
}

class AadAppSetCommand extends GraphCommand {
  private static aadApplicationPlatform: string[] = ['spa', 'web', 'publicClient'];

  public get name(): string {
    return commands.APP_SET;
  }

  public get description(): string {
    return 'Updates Azure AD app registration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.objectId = typeof args.options.objectId !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.platform = typeof args.options.platform !== 'undefined';
    telemetryProps.redirectUris = typeof args.options.redirectUris !== 'undefined';
    telemetryProps.redirectUrisToRemove = typeof args.options.redirectUrisToRemove !== 'undefined';
    telemetryProps.uri = typeof args.options.uri !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAppObjectId(args, logger)
      .then(objectId => this.configureUri(args, objectId, logger))
      .then(objectId => this.configureRedirectUris(args, objectId, logger))
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private getAppObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.objectId) {
      return Promise.resolve(args.options.objectId);
    }

    const { appId, name } = args.options;

    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Azure AD app ${appId ? appId : name}...`);
    }

    const filter: string = appId ?
      `appId eq '${encodeURIComponent(appId)}'` :
      `displayName eq '${encodeURIComponent(name as string)}'`;

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=${filter}&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { id: string }[] }>(requestOptions)
      .then((res: { value: { id: string }[] }): Promise<string> => {
        if (res.value.length === 1) {
          return Promise.resolve(res.value[0].id);
        }

        if (res.value.length === 0) {
          const applicationIdentifier = appId ? `ID ${appId}` : `name ${name}`;
          return Promise.reject(`No Azure AD application registration with ${applicationIdentifier} found`);
        }

        return Promise.reject(`Multiple Azure AD application registration with name ${name} found. Please disambiguate (app object IDs): ${res.value.map(a => a.id).join(', ')}`);
      });
  }

  private configureUri(args: CommandArgs, objectId: string, logger: Logger): Promise<string> {
    if (!args.options.uri) {
      return Promise.resolve(objectId);
    }

    if (this.verbose) {
      logger.logToStderr(`Configuring Azure AD application ID URI...`);
    }

    const applicationInfo: any = {
      identifierUris: [args.options.uri]
    };

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: applicationInfo
    };

    return request
      .patch(requestOptions)
      .then(_ => Promise.resolve(objectId));
  }

  private configureRedirectUris(args: CommandArgs, objectId: string, logger: Logger): Promise<string> {
    if (!args.options.redirectUris && !args.options.redirectUrisToRemove) {
      return Promise.resolve(objectId);
    }

    if (this.verbose) {
      logger.logToStderr(`Configuring Azure AD application redirect URIs...`);
    }

    const getAppRequestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<Application>(getAppRequestOptions)
      .then((application: Application): Promise<void> => {
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

        const requestOptions: AxiosRequestConfig = {
          url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
          headers: {
            'content-type': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: applicationPatch
        };

        return request.patch(requestOptions);
      })
      .then(_ => Promise.resolve(objectId));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--appId [appId]' },
      { option: '--objectId [objectId]' },
      { option: '-n, --name [name]' },
      { option: '-u, --uri [uri]' },
      { option: '-r, --redirectUris [redirectUris]' },
      {
        option: '--platform [platform]',
        autocomplete: AadAppSetCommand.aadApplicationPlatform
      },
      { option: '--redirectUrisToRemove [redirectUrisToRemove]' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.appId &&
      !args.options.objectId &&
      !args.options.name) {
      return 'Specify either appId, objectId or name';
    }

    if ((args.options.appId && args.options.objectId) ||
      (args.options.appId && args.options.name) ||
      (args.options.objectId && args.options.name)) {
      return 'Specify either appId, objectId or name but not both';
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
}

module.exports = new AadAppSetCommand();