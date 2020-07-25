import config from '../../../../config';
import commands from '../../commands';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CommandOption, CommandValidate } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo } from '../../spo';
import { SpoOperation } from './SpoOperation';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  url: string;
  title?: string;
  sharing?: string;
  resourceQuota?: string | number;
  resourceQuotaWarningLevel?: string | number;
  storageQuota?: string | number;
  storageQuotaWarningLevel?: string | number;
  allowSelfServiceUpgrade?: string;
  owners?: string;
  lockState?: string;
  noScriptSite?: string;
  wait: boolean;
}

class SpoSiteClassicSetCommand extends SpoCommand {
  private context?: FormDigestInfo;
  private spoAdminUrl?: string;
  private tenantId?: string;
  private dots?: string;
  private timeout?: NodeJS.Timer;

  public get name(): string {
    return commands.SITE_CLASSIC_SET;
  }

  public get description(): string {
    return 'Change classic site settings';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.sharing = args.options.sharing;
    telemetryProps.resourceQuota = typeof args.options.resourceQuota !== 'undefined';
    telemetryProps.resourceQuotaWarningLevel = typeof args.options.resourceQuotaWarningLevel !== 'undefined';
    telemetryProps.storageQuota = typeof args.options.storageQuota !== 'undefined';
    telemetryProps.storageQuotaWarningLevel = typeof args.options.storageQuotaWarningLevel !== 'undefined';
    telemetryProps.allowSelfServiceUpgrade = args.options.allowSelfServiceUpgrade;
    telemetryProps.owners = typeof args.options.owners !== 'undefined';
    telemetryProps.lockState = args.options.lockState;
    telemetryProps.noScriptSite = args.options.noScriptSite;
    telemetryProps.wait = args.options.wait;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this.dots = '';

    this
      .getTenantId(cmd, this.debug)
      .then((_tenantId: string): Promise<string> => {
        this.tenantId = _tenantId;

        return this.getSpoAdminUrl(cmd, this.debug)
      })
      .then((_spoAdminUrl: string): Promise<FormDigestInfo> => {
        this.spoAdminUrl = _spoAdminUrl;

        return this.ensureFormDigest(this.spoAdminUrl, cmd, this.context, this.debug);
      })
      .then((res: FormDigestInfo): Promise<string> => {
        this.context = res;
        if (this.verbose) {
          cmd.log(`Setting basic properties ${args.options.url}...`);
        }

        const basicProperties: string[] = [
          'title',
          'sharing',
          'resourceQuota',
          'resourceQuotaWarningLevel',
          'storageQuota',
          'storageQuotaWarningLevel',
          'allowSelfServiceUpgrade',
          'noScriptSite'
        ];

        let updateBasicProperties: boolean = false;
        for (let i: number = 0; i < basicProperties.length; i++) {
          if (typeof (args.options as any)[basicProperties[i]] !== 'undefined') {
            updateBasicProperties = true;
            break;
          }
        }

        if (!updateBasicProperties) {
          return Promise.resolve(undefined as any);
        }

        let i: number = 0;
        const updates: string[] = [];

        if (args.options.title) {
          updates.push(`<SetProperty Id="${++i}" ObjectPathId="5" Name="Title"><Parameter Type="String">${Utils.escapeXml(args.options.title)}</Parameter></SetProperty>`);
        }
        if (args.options.sharing) {
          const sharing: number = ['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly'].indexOf(args.options.sharing);
          updates.push(`<SetProperty Id="${++i}" ObjectPathId="5" Name="SharingCapability"><Parameter Type="Enum">${sharing}</Parameter></SetProperty>`);
        }
        if (args.options.resourceQuota) {
          updates.push(`<SetProperty Id="${++i}" ObjectPathId="5" Name="UserCodeMaximumLevel"><Parameter Type="Double">${args.options.resourceQuota}</Parameter></SetProperty>`);
        }
        if (args.options.resourceQuotaWarningLevel) {
          updates.push(`<SetProperty Id="${++i}" ObjectPathId="5" Name="UserCodeWarningLevel"><Parameter Type="Double">${args.options.resourceQuotaWarningLevel}</Parameter></SetProperty>`);
        }
        if (args.options.storageQuota) {
          updates.push(`<SetProperty Id="${++i}" ObjectPathId="5" Name="StorageMaximumLevel"><Parameter Type="Int64">${args.options.storageQuota}</Parameter></SetProperty>`);
        }
        if (args.options.storageQuotaWarningLevel) {
          updates.push(`<SetProperty Id="${++i}" ObjectPathId="5" Name="StorageWarningLevel"><Parameter Type="Int64">${args.options.storageQuotaWarningLevel}</Parameter></SetProperty>`);
        }
        if (typeof args.options.allowSelfServiceUpgrade !== 'undefined') {
          updates.push(`<SetProperty Id="${++i}" ObjectPathId="5" Name="AllowSelfServiceUpgrade"><Parameter Type="Boolean">${args.options.allowSelfServiceUpgrade}</Parameter></SetProperty>`);
        }
        if (typeof args.options.noScriptSite !== 'undefined') {
          const noScriptSite: number = args.options.noScriptSite === 'true' ? 2 : 1;
          updates.push(`<SetProperty Id="${++i}" ObjectPathId="5" Name="DenyAddAndCustomizePages"><Parameter Type="Enum">${noScriptSite}</Parameter></SetProperty>`);
        }

        const pos: number = (this.tenantId as string).indexOf('|') + 1;

        const requestOptions: any = {
          url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': this.context.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${updates.join('')}<ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|${(this.tenantId as string).substr(pos, (this.tenantId as string).indexOf('&') - pos)}&#xA;SiteProperties&#xA;${encodeURIComponent(args.options.url)}" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res?: string): Promise<void> => {
        return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
          if (!res) {
            resolve();
            return;
          }

          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            reject(response.ErrorInfo.ErrorMessage);
          }
          else {
            const operation: SpoOperation = json[json.length - 1];
            let isComplete: boolean = operation.IsComplete;
            if (!args.options.wait || isComplete) {
              resolve();
              return;
            }

            this.timeout = setTimeout(() => {
              this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), this.spoAdminUrl as string, resolve, reject, cmd, this.context as FormDigestInfo, this.dots, this.timeout);
            }, operation.PollingInterval);
          }
        });
      })
      .then((): Promise<FormDigestInfo> => {
        return this.ensureFormDigest(this.spoAdminUrl as string, cmd, this.context, this.debug);
      })
      .then((res: FormDigestInfo): Promise<void> => {
        this.context = res;
        return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
          if (!args.options.owners) {
            resolve();
            return;
          }

          Promise.all(args.options.owners.split(',').map(o => {
            return this.setAdmin(cmd, args.options.url, o.trim());
          }))
            .then((): void => {
              resolve();
            }, (err: any): void => {
              reject(err);
            });
        });
      })
      .then((): Promise<string> => {
        if (!args.options.lockState) {
          return Promise.resolve(undefined as any);
        }

        const requestOptions: any = {
          url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': (this.context as FormDigestInfo).FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">${Utils.escapeXml(args.options.lockState)}</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res?: string): Promise<void> => {
        return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
          if (!res) {
            resolve();
            return;
          }

          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            reject(response.ErrorInfo.ErrorMessage);
          }
          else {
            const operation: SpoOperation = json[json.length - 1];
            let isComplete: boolean = operation.IsComplete;
            if (!args.options.wait || isComplete) {
              resolve();
              return;
            }

            this.timeout = setTimeout(() => {
              this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), this.spoAdminUrl as string, resolve, reject, cmd, this.context as FormDigestInfo, this.dots, this.timeout);
            }, operation.PollingInterval);
          }
        });
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  private setAdmin(cmd: CommandInstance, siteUrl: string, principal: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .ensureFormDigest(this.spoAdminUrl as string, cmd, this.context, this.debug)
        .then((res: FormDigestInfo): Promise<string> => {
          this.context = res;
          const requestOptions: any = {
            url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': this.context.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">${Utils.escapeXml(siteUrl)}</Parameter><Parameter Type="String">${Utils.escapeXml(principal)}</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The absolute site url'
      },
      {
        option: '-t, --title [title]',
        description: 'The site title'
      },
      {
        option: '--sharing [sharing]',
        description: 'Sharing capabilities for the site. Allowed values: Disabled|ExternalUserSharingOnly|ExternalUserAndGuestSharing|ExistingExternalUserSharingOnly',
        autocomplete: ['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly']
      },
      {
        option: '--resourceQuota [resourceQuota]',
        description: 'The quota for this site collection in Sandboxed Solutions units'
      },
      {
        option: '--resourceQuotaWarningLevel [resourceQuotaWarningLevel]',
        description: 'The warning level for the resource quota'
      },
      {
        option: '--storageQuota [storageQuota]',
        description: 'The storage quota for this site collection in megabytes'
      },
      {
        option: '--storageQuotaWarningLevel [storageQuotaWarningLevel]',
        description: 'The warning level for the storage quota in megabytes'
      },
      {
        option: '--allowSelfServiceUpgrade [allowSelfServiceUpgrade]',
        description: 'Set to allow tenant administrators to upgrade the site collection'
      },
      {
        option: '--owners [owners]',
        description: 'Comma-separated list of users to add as site collection administrators'
      },
      {
        option: '--lockState [lockState]',
        description: 'Sets site\'s lock state. Allowed values Unlock|NoAdditions|ReadOnly|NoAccess',
        autocomplete: ['Unlock', 'NoAdditions', 'ReadOnly', 'NoAccess']
      },
      {
        option: '--noScriptSite [noScriptSite]',
        description: 'Specifies if the site allows custom script or not'
      },
      {
        option: '--wait',
        description: 'Wait for the settings to be applied before completing the command'
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

      if (args.options.sharing &&
        ['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly'].indexOf(args.options.sharing) === -1) {
        return `${args.options.sharing} is not a valid value for the sharing option. Allowed values Disabled|ExternalUserSharingOnly|ExternalUserAndGuestSharing|ExistingExternalUserSharingOnly`;
      }

      if (args.options.resourceQuota &&
        typeof args.options.resourceQuota !== 'number') {
        return `${args.options.resourceQuota} is not a number`;
      }

      if (args.options.resourceQuotaWarningLevel &&
        typeof args.options.resourceQuotaWarningLevel !== 'number') {
        return `${args.options.resourceQuotaWarningLevel} is not a number`;
      }

      if (args.options.resourceQuota &&
        args.options.resourceQuotaWarningLevel &&
        args.options.resourceQuotaWarningLevel > args.options.resourceQuota) {
        return `resourceQuotaWarningLevel must not exceed the resourceQuota`;
      }

      if (args.options.storageQuota &&
        typeof args.options.storageQuota !== 'number') {
        return `${args.options.storageQuota} is not a number`;
      }

      if (args.options.storageQuotaWarningLevel &&
        typeof args.options.storageQuotaWarningLevel !== 'number') {
        return `${args.options.storageQuotaWarningLevel} is not a number`;
      }

      if (args.options.storageQuota &&
        args.options.storageQuotaWarningLevel &&
        args.options.storageQuotaWarningLevel > args.options.storageQuota) {
        return `storageQuotaWarningLevel must not exceed the storageQuota`;
      }

      if (args.options.allowSelfServiceUpgrade &&
        args.options.allowSelfServiceUpgrade !== 'true' &&
        args.options.allowSelfServiceUpgrade !== 'false') {
        return `${args.options.allowSelfServiceUpgrade} is not a valid boolean value`;
      }

      if (args.options.lockState &&
        ['Unlock', 'NoAdditions', 'ReadOnly', 'NoAccess'].indexOf(args.options.lockState) === -1) {
        return `${args.options.lockState} is not a valid value for the lockState option. Allowed values Unlock|NoAdditions|ReadOnly|NoAccess`;
      }

      if (args.options.noScriptSite &&
        args.options.noScriptSite !== 'true' &&
        args.options.noScriptSite !== 'false') {
        return `${args.options.noScriptSite} is not a valid boolean value`;
      }

      return true;
    };
  }
}

module.exports = new SpoSiteClassicSetCommand();
