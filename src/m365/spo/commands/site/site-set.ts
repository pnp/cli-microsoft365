import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import Command, {
  CommandOption,
  CommandValidate,
  CommandTypes,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import * as spoSiteClassicSetCommand from './site-classic-set';
import { Options as SpoSiteClassicSetCommandOptions } from './site-classic-set';
import * as aadO365GroupSetCommand from '../../../aad/commands/o365group/o365group-set';
import { Options as AadO365GroupSetCommandOptions } from '../../../aad/commands/o365group/o365group-set';
import * as spoSiteDesignApplyCommand from '../sitedesign/sitedesign-apply';
import { Options as SpoSiteDesignApplyCommandOptions } from '../sitedesign/sitedesign-apply';
import { SharingCapabilities } from '../site/SharingCapabilities';
import * as chalk from 'chalk';
import { CommandInstance, Cli } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  classification?: string;
  disableFlows?: string;
  isPublic?: string;
  owners?: string;
  shareByEmailEnabled?: string;
  siteDesignId?: string;
  title?: string;
  url: string;
  sharingCapability?: string;
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
    telemetryProps.siteDesignId = typeof args.options.siteDesignId !== undefined;
    telemetryProps.sharingCapabilities = args.options.sharingCapability;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .loadSiteIds(args.options.url, cmd)
      .then((): Promise<void> => {
        if (this.groupId === '00000000-0000-0000-0000-000000000000') {
          if (this.debug) {
            cmd.log('Site is not groupified');
          }

          return this.updateSite(cmd, args);
        }
        else {
          if (this.debug) {
            cmd.log(`Site attached to group ${this.groupId}`);
          }

          return this.updateGroupifiedSite(cmd, args);
        }
      })
      .then((): Promise<void> => this.updateSharedProperties(cmd, args))
      .then((): Promise<void> => this.applySiteDesign(cmd, args))
      .then((): Promise<void> => this.setSharingCapabilities(cmd, args))
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => {
        if (err instanceof CommandError) {
          err = (err as CommandError).message;
        }

        this.handleRejectedPromise(err, cmd, cb)
      });
  }

  private updateSite(cmd: CommandInstance, args: CommandArgs): Promise<void> {
    if (typeof args.options.isPublic !== 'undefined') {
      return Promise.reject(`The isPublic option can't be set on a site that is not groupified`);
    }

    if (!args.options.title &&
      !args.options.owners) {
      return Promise.resolve();
    }

    const options: SpoSiteClassicSetCommandOptions = {
      url: args.options.url,
      title: args.options.title,
      owners: args.options.owners,
      wait: true,
      debug: this.debug,
      verbose: this.verbose
    };
    return Cli.executeCommand((spoSiteClassicSetCommand as Command).name, spoSiteClassicSetCommand as Command, { options: { ...options, _: [] } });
  }

  private updateGroupifiedSite(cmd: CommandInstance, args: CommandArgs): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (typeof args.options.title === 'undefined' &&
        typeof args.options.isPublic === 'undefined' &&
        typeof args.options.owners === 'undefined') {
        return resolve();
      }

      let spoAdminUrl: string;

      const promises: Promise<void>[] = [];

      if (typeof args.options.title !== 'undefined') {
        promises.push(this
          .getSpoAdminUrl(cmd, this.debug)
          .then((_spoAdminUrl: string): Promise<FormDigestInfo> => {
            spoAdminUrl = _spoAdminUrl;

            return this.getRequestDigest(spoAdminUrl);
          })
          .then((formDigest: FormDigestInfo) => {
            const requestOptions: any = {
              url: `${spoAdminUrl}/_api/SPOGroup/UpdateGroupPropertiesBySiteId`,
              headers: {
                accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;charset=utf-8',
                'X-RequestDigest': formDigest.FormDigestValue
              },
              body: {
                groupId: this.groupId,
                siteId: this.siteId,
                displayName: args.options.title
              },
              json: true
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
        promises.push(Cli.executeCommand((aadO365GroupSetCommand as Command).name, aadO365GroupSetCommand as Command, { options: { ...commandOptions, _: [] } }));
      }

      promises.push(this.setGroupifiedSiteOwners(cmd, args));

      Promise
        .all(promises)
        .then((): void => {
          resolve();
        }, (error: any): void => {
          reject(error);
        })
    });
  }

  private setGroupifiedSiteOwners(cmd: CommandInstance, args: CommandArgs): Promise<void> {
    if (typeof args.options.owners === 'undefined') {
      return Promise.resolve();
    }

    const owners: string[] = args.options.owners.split(',').map(o => o.trim());

    if (this.verbose) {
      cmd.log('Retrieving user information to set group owners...');
    }

    let spoAdminUrl: string;

    return this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<{ value: { id: string; }[] }> => {
        spoAdminUrl = _spoAdminUrl;

        const requestOptions: any = {
          url: `https://graph.microsoft.com/v1.0/users?$filter=${owners.map(o => `userPrincipalName eq '${o}'`).join(' or ')}&$select=id`,
          headers: {
            'content-type': 'application/json;odata.metadata=none'
          },
          json: true
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

  private updateSharedProperties(cmd: CommandInstance, args: CommandArgs): Promise<void> {
    if (typeof args.options.classification === 'undefined' &&
      typeof args.options.disableFlows === 'undefined' &&
      typeof args.options.shareByEmailEnabled === 'undefined') {
      return Promise.resolve();
    }

    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (this.verbose) {
        cmd.log(`Retrieving request digest...`);
      }

      this
        .getRequestDigest(args.options.url)
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            cmd.log(`Updating site ${args.options.url} properties...`);
          }

          let propertyId: number = 27;
          const payload: string[] = [];
          if (typeof args.options.classification === 'string') {
            payload.push(`<SetProperty Id="${propertyId++}" ObjectPathId="5" Name="Classification"><Parameter Type="String">${Utils.escapeXml(args.options.classification)}</Parameter></SetProperty>`);
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
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${payload.join('')}</Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}" /></ObjectPaths></Request>`
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

  private applySiteDesign(cmd: CommandInstance, args: CommandArgs): Promise<void> {
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
    return Cli.executeCommand((spoSiteDesignApplyCommand as Command).name, spoSiteDesignApplyCommand as Command, { options: { ...options, _: [] } });
  }

  private setSharingCapabilities(cmd: CommandInstance, args: CommandArgs): Promise<void> {
    if (typeof args.options.sharingCapability === 'undefined') {
      return Promise.resolve();
    }

    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (this.verbose) {
        cmd.log(`Retrieving request digest...`);
      }

      const sharingCapability: SharingCapabilities = SharingCapabilities[(args.options.sharingCapability as keyof typeof SharingCapabilities)];

      this
        .getSpoAdminUrl(cmd, this.debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          this.spoAdminUrl = _spoAdminUrl;

          return this.getRequestDigest(this.spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            cmd.log(`Setting sharing for site  ${args.options.url} as ${args.options.sharingCapability}`);
          }

          const requestOptions: any = {
            url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1"/><ObjectPath Id="4" ObjectPathId="3"/><SetProperty Id="5" ObjectPathId="3" Name="SharingCapability"><Parameter Type="Enum">${sharingCapability}</Parameter></SetProperty><ObjectPath Id="7" ObjectPathId="6"/><ObjectIdentityQuery Id="8" ObjectPathId="3"/></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="6" ParentId="3" Name="Update"/></ObjectPaths></Request>`
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

  private loadSiteIds(siteUrl: string, cmd: CommandInstance): Promise<void> {
    if (this.debug) {
      cmd.log('Loading site IDs...');
    }

    const requestOptions: any = {
      url: `${siteUrl}/_api/site?$select=GroupId,Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    return request
      .get<{ GroupId: string; Id: string }>(requestOptions)
      .then((siteInfo: { GroupId: string; Id: string }): Promise<void> => {
        this.groupId = siteInfo.GroupId;
        this.siteId = siteInfo.Id;

        if (this.debug) {
          cmd.log(`Retrieved site IDs. siteId: ${this.siteId}, groupId: ${this.groupId}`);
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

    for (let sharingCapability in SharingCapabilities) {
      if (typeof SharingCapabilities[sharingCapability] === 'number') {
        result.push(sharingCapability);
      }
    }

    return result;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The URL of the site collection to update'
      },
      {
        option: '-i, --id [id]',
        description: 'The ID of the site collection to update (deprecated; id is automatically retrieved and does not need to be specified)'
      },
      {
        option: '--classification [classification]',
        description: 'The new classification for the site collection'
      },
      {
        option: '--disableFlows [disableFlows]',
        description: 'Set to true to disable using Microsoft Flow in this site collection'
      },
      {
        option: '--isPublic [isPublic]',
        description: 'Set to true to make the group linked to the site public or to false to make it private'
      },
      {
        option: '--owners [owners]',
        description: 'Comma-separated list of users to add as site collection administrators'
      },
      {
        option: '--shareByEmailEnabled [shareByEmailEnabled]',
        description: 'Set to true to allow to share files with guests and to false to disallow it'
      },
      {
        option: '--siteDesignId [siteDesignId]',
        description: 'Id of the custom site design to apply to the site'
      },
      {
        option: '--title [title]',
        description: 'The new title for the site collection'
      },
      {
        option: '--sharingCapability [sharingCapability]',
        description: `The sharing capability for the Site. Allowed values ${this.sharingCapabilities.join('|')}.`,
        autocomplete: this.sharingCapabilities
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (typeof args.options.classification === 'undefined' &&
        typeof args.options.disableFlows === 'undefined' &&
        typeof args.options.title === 'undefined' &&
        typeof args.options.isPublic === 'undefined' &&
        typeof args.options.owners === 'undefined' &&
        typeof args.options.shareByEmailEnabled === 'undefined' &&
        typeof args.options.siteDesignId === 'undefined' &&
        typeof args.options.sharingCapability === 'undefined') {
        return 'Specify at least one property to update';
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
        if (!Utils.isValidGuid(args.options.siteDesignId)) {
          return `${args.options.siteDesignId} is not a valid GUID`;
        }
      }

      if (args.options.sharingCapability &&
        this.sharingCapabilities.indexOf(args.options.sharingCapability) < 0) {
        return `${args.options.sharingCapability} is not a valid value for the sharingCapability option. Allowed values are ${this.sharingCapabilities.join('|')}`;
      }

      return true;
    };
  }

  public types(): CommandTypes {
    // required to support passing empty strings as valid values
    return {
      string: ['classification']
    }
  }
}

module.exports = new SpoSiteSetCommand();