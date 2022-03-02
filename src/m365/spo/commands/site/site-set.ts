import { Cli, Logger } from '../../../../cli';
import Command, {
  CommandError, CommandOption,
  CommandTypes
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, FormDigestInfo, spo, urlUtil, validation } from '../../../../utils';
import * as aadO365GroupSetCommand from '../../../aad/commands/o365group/o365group-set';
import { Options as AadO365GroupSetCommandOptions } from '../../../aad/commands/o365group/o365group-set';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SharingCapabilities } from '../site/SharingCapabilities';
import * as spoSiteDesignApplyCommand from '../sitedesign/sitedesign-apply';
import { Options as SpoSiteDesignApplyCommandOptions } from '../sitedesign/sitedesign-apply';
import * as spoSiteClassicSetCommand from './site-classic-set';
import { Options as SpoSiteClassicSetCommandOptions } from './site-classic-set';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  classification?: string;
  disableFlows?: string;
  isPublic?: string;
  owners?: string;
  shareByEmailEnabled?: string;
  siteDesignId?: string;
  title?: string;
  description?: string;
  url: string;
  sharingCapability?: string;
  siteLogoUrl?: string;
}

class SpoSiteSetCommand extends SpoCommand {
  private groupId: string | undefined;
  private siteId: string | undefined;
  private spoAdminUrl?: string;

  public get name(): string {
    return commands.SITE_SET;
  }

  public get description(): string {
    return 'Updates properties of the specified site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.classification = typeof args.options.classification === 'string';
    telemetryProps.disableFlows = args.options.disableFlows;
    telemetryProps.isPublic = args.options.isPublic;
    telemetryProps.owners = typeof args.options.owners !== 'undefined';
    telemetryProps.shareByEmailEnabled = args.options.shareByEmailEnabled;
    telemetryProps.title = typeof args.options.title === 'string';
    telemetryProps.description = typeof args.options.description === 'string';
    telemetryProps.siteDesignId = typeof args.options.siteDesignId !== undefined;
    telemetryProps.sharingCapabilities = args.options.sharingCapability;
    telemetryProps.siteLogoUrl = typeof args.options.siteLogoUrl !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .loadSiteIds(args.options.url, logger)
      .then((): Promise<void> => {
        if (this.groupId === '00000000-0000-0000-0000-000000000000') {
          if (this.debug) {
            logger.logToStderr('Site is not groupified');
          }

          return this.updateSite(logger, args);
        }
        else {
          if (this.debug) {
            logger.logToStderr(`Site attached to group ${this.groupId}`);
          }

          return this.updateGroupifiedSite(logger, args);
        }
      })
      .then((): Promise<void> => this.updateSharedProperties(logger, args))
      .then((): Promise<void> => this.applySiteDesign(logger, args))
      .then((): Promise<void> => this.setSharingCapabilities(logger, args))
      .then((): Promise<void> => this.setLogo(logger, args))
      .then(_ => cb(), (err: any): void => {
        if (err instanceof CommandError) {
          err = (err as CommandError).message;
        }

        this.handleRejectedPromise(err, logger, cb);
      });
  }

  private setLogo(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.siteLogoUrl === 'undefined') {
      return Promise.resolve();
    }

    if (this.debug) {
      logger.logToStderr(`Setting the site its logo...`);
    }

    const logoUrl = args.options.siteLogoUrl ? urlUtil.getServerRelativePath(args.options.url, args.options.siteLogoUrl) : "";

    const requestOptions: any = {
      url: `${args.options.url}/_api/siteiconmanager/setsitelogo`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: {
        relativeLogoUrl: logoUrl
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private updateSite(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.isPublic !== 'undefined') {
      return Promise.reject(`The isPublic option can't be set on a site that is not groupified`);
    }

    if (!args.options.title &&
      !args.options.owners &&
      !args.options.description) {
      return Promise.resolve();
    }

    const options: SpoSiteClassicSetCommandOptions = {
      url: args.options.url,
      title: args.options.title,
      description: args.options.description,
      owners: args.options.owners,
      wait: true,
      debug: this.debug,
      verbose: this.verbose
    };

    return Cli.executeCommand(spoSiteClassicSetCommand as Command, { options: { ...options, _: [] } });
  }

  private updateGroupifiedSite(logger: Logger, args: CommandArgs): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (typeof args.options.title === 'undefined' &&
        typeof args.options.description === 'undefined' &&
        typeof args.options.isPublic === 'undefined' &&
        typeof args.options.owners === 'undefined') {
        return resolve();
      }

      let spoAdminUrl: string;

      const promises: Promise<void>[] = [];

      if (typeof args.options.title !== 'undefined') {
        promises.push(spo
          .getSpoAdminUrl(logger, this.debug)
          .then((_spoAdminUrl: string): Promise<FormDigestInfo> => {
            spoAdminUrl = _spoAdminUrl;

            return spo.getRequestDigest(spoAdminUrl);
          })
          .then((formDigest: FormDigestInfo) => {
            const requestOptions: any = {
              url: `${spoAdminUrl}/_api/SPOGroup/UpdateGroupPropertiesBySiteId`,
              headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;charset=utf-8',
                'X-RequestDigest': formDigest.FormDigestValue
              },
              data: {
                groupId: this.groupId,
                siteId: this.siteId,
                displayName: args.options.title
              },
              responseType: 'json'
            };
            return request.post(requestOptions);
          }));
      }

      if (typeof args.options.isPublic !== 'undefined') {
        const commandOptions: AadO365GroupSetCommandOptions = {
          id: this.groupId as string,
          isPrivate: (args.options.isPublic === 'false').toString(),
          debug: this.debug,
          verbose: this.verbose
        };
        promises.push(Cli.executeCommand(aadO365GroupSetCommand as Command, { options: { ...commandOptions, _: [] } }));
      }

      if (args.options.description) {
        promises.push(this.setGroupifiedSiteDescription(args.options.description));
      }

      promises.push(this.setGroupifiedSiteOwners(logger, args));

      Promise
        .all(promises)
        .then((): void => {
          resolve();
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private setGroupifiedSiteDescription(description: string): Promise<void> {
    const requestOptions: any = {
      url: `https://graph.microsoft.com/v1.0/groups/${this.groupId}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      data: {
        description: description
      }
    };

    return request.patch(requestOptions);
  }

  private setGroupifiedSiteOwners(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.owners === 'undefined') {
      return Promise.resolve();
    }

    const owners: string[] = args.options.owners.split(',').map(o => o.trim());

    if (this.verbose) {
      logger.logToStderr('Retrieving user information to set group owners...');
    }

    let spoAdminUrl: string;

    return spo
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<{ value: { id: string; }[] }> => {
        spoAdminUrl = _spoAdminUrl;

        const requestOptions: any = {
          url: `https://graph.microsoft.com/v1.0/users?$filter=${owners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
          headers: {
            'content-type': 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: { value: { id: string; }[] }): Promise<any> => {
        if (res.value.length === 0) {
          return Promise.resolve();
        }

        return Promise.all(res.value.map(user => {
          const requestOptions: any = {
            url: `${spoAdminUrl}/_api/SP.Directory.DirectorySession/Group('${this.groupId}')/Owners/Add(objectId='${user.id}', principalName='')`,
            headers: {
              'content-type': 'application/json;odata=verbose'
            }
          };

          return request.post(requestOptions);
        }));
      });
  }

  private updateSharedProperties(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.classification === 'undefined' &&
      typeof args.options.disableFlows === 'undefined' &&
      typeof args.options.shareByEmailEnabled === 'undefined') {
      return Promise.resolve();
    }

    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (this.verbose) {
        logger.logToStderr(`Retrieving request digest...`);
      }

      spo
        .getRequestDigest(args.options.url)
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            logger.logToStderr(`Updating site ${args.options.url} properties...`);
          }

          let propertyId: number = 27;
          const payload: string[] = [];
          if (typeof args.options.classification === 'string') {
            payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Classification"><Parameter Type="String">${formatting.escapeXml(args.options.classification)}</Parameter></SetProperty>`);
          }
          if (typeof args.options.disableFlows === 'string') {
            payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">${args.options.disableFlows === 'true'}</Parameter></SetProperty>`);
          }
          if (typeof args.options.shareByEmailEnabled === 'string') {
            payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">${args.options.shareByEmailEnabled === 'true'}</Parameter></SetProperty>`);
          }

          // update site via the Site object
          const requestOptions: any = {
            url: `${args.options.url}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${payload.join('')}</Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}" /></ObjectPaths></Request>`
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            reject(response.ErrorInfo.ErrorMessage);
          }
          else {
            resolve();
          }
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  private applySiteDesign(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.siteDesignId === 'undefined') {
      return Promise.resolve();
    }

    const options: SpoSiteDesignApplyCommandOptions = {
      webUrl: args.options.url,
      id: args.options.siteDesignId,
      asTask: false,
      debug: this.debug,
      verbose: this.verbose
    };
    return Cli.executeCommand(spoSiteDesignApplyCommand as Command, { options: { ...options, _: [] } });
  }

  private setSharingCapabilities(logger: Logger, args: CommandArgs): Promise<void> {
    if (typeof args.options.sharingCapability === 'undefined') {
      return Promise.resolve();
    }

    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (this.verbose) {
        logger.logToStderr(`Retrieving request digest...`);
      }

      const sharingCapability: SharingCapabilities = SharingCapabilities[(args.options.sharingCapability as keyof typeof SharingCapabilities)];

      spo
        .getSpoAdminUrl(logger, this.debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          this.spoAdminUrl = _spoAdminUrl;

          return spo.getRequestDigest(this.spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            logger.logToStderr(`Setting sharing for site  ${args.options.url} as ${args.options.sharingCapability}`);
          }

          const requestOptions: any = {
            url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1"/><ObjectPath Id="4" ObjectPathId="3"/><SetProperty Id="5" ObjectPathId="3" Name="SharingCapability"><Parameter Type="Enum">${sharingCapability}</Parameter></SetProperty><ObjectPath Id="7" ObjectPathId="6"/><ObjectIdentityQuery Id="8" ObjectPathId="3"/></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="6" ParentId="3" Name="Update"/></ObjectPaths></Request>`
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            reject(response.ErrorInfo.ErrorMessage);
          }
          else {
            resolve();
          }
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  private loadSiteIds(siteUrl: string, logger: Logger): Promise<void> {
    if (this.debug) {
      logger.logToStderr('Loading site IDs...');
    }

    const requestOptions: any = {
      url: `${siteUrl}/_api/site?$select=GroupId,Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get<{ GroupId: string; Id: string }>(requestOptions)
      .then((siteInfo: { GroupId: string; Id: string }): Promise<void> => {
        this.groupId = siteInfo.GroupId;
        this.siteId = siteInfo.Id;

        if (this.debug) {
          logger.logToStderr(`Retrieved site IDs. siteId: ${this.siteId}, groupId: ${this.groupId}`);
        }

        return Promise.resolve();
      });
  }

  /**
   * Maps the base sharingCapability enum to string array so it can 
   * more easily be used in validation or descriptions.
   */
  protected get sharingCapabilities(): string[] {
    const result: string[] = [];

    for (const sharingCapability in SharingCapabilities) {
      if (typeof SharingCapabilities[sharingCapability] === 'number') {
        result.push(sharingCapability);
      }
    }

    return result;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--classification [classification]'
      },
      {
        option: '--disableFlows [disableFlows]'
      },
      {
        option: '--isPublic [isPublic]'
      },
      {
        option: '--owners [owners]'
      },
      {
        option: '--shareByEmailEnabled [shareByEmailEnabled]'
      },
      {
        option: '--siteDesignId [siteDesignId]'
      },
      {
        option: '--title [title]'
      },
      {
        option: '--description [description]'
      },
      {
        option: '--siteLogoUrl [siteLogoUrl]'
      },
      {
        option: '--sharingCapability [sharingCapability]',
        autocomplete: this.sharingCapabilities
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.url);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (typeof args.options.classification === 'undefined' &&
      typeof args.options.disableFlows === 'undefined' &&
      typeof args.options.title === 'undefined' &&
      typeof args.options.description === 'undefined' &&
      typeof args.options.isPublic === 'undefined' &&
      typeof args.options.owners === 'undefined' &&
      typeof args.options.shareByEmailEnabled === 'undefined' &&
      typeof args.options.siteDesignId === 'undefined' &&
      typeof args.options.sharingCapability === 'undefined' &&
      typeof args.options.siteLogoUrl === 'undefined') {
      return 'Specify at least one property to update';
    }

    if (typeof args.options.siteLogoUrl !== 'undefined' && typeof args.options.siteLogoUrl !== 'string') {
      return `${args.options.siteLogoUrl} is not a valid value for the siteLogoUrl option. Specify the logo URL or an empty string "" to unset the logo.`;
    }

    if (typeof args.options.disableFlows === 'string' &&
      args.options.disableFlows !== 'true' &&
      args.options.disableFlows !== 'false') {
      return `${args.options.disableFlows} is not a valid value for the disableFlow option. Allowed values are true|false`;
    }

    if (typeof args.options.isPublic === 'string' &&
      args.options.isPublic !== 'true' &&
      args.options.isPublic !== 'false') {
      return `${args.options.isPublic} is not a valid value for the isPublic option. Allowed values are true|false`;
    }

    if (typeof args.options.shareByEmailEnabled === 'string' &&
      args.options.shareByEmailEnabled !== 'true' &&
      args.options.shareByEmailEnabled !== 'false') {
      return `${args.options.shareByEmailEnabled} is not a valid value for the shareByEmailEnabled option. Allowed values are true|false`;
    }

    if (args.options.siteDesignId) {
      if (!validation.isValidGuid(args.options.siteDesignId)) {
        return `${args.options.siteDesignId} is not a valid GUID`;
      }
    }

    if (args.options.sharingCapability &&
      this.sharingCapabilities.indexOf(args.options.sharingCapability) < 0) {
      return `${args.options.sharingCapability} is not a valid value for the sharingCapability option. Allowed values are ${this.sharingCapabilities.join('|')}`;
    }

    return true;
  }

  public types(): CommandTypes {
    // required to support passing empty strings as valid values
    return {
      string: ['classification']
    };
  }
}

module.exports = new SpoSiteSetCommand();