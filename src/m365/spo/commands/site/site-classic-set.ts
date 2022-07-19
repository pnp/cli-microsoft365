import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, formatting, FormDigestInfo, spo, SpoOperation, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  url: string;
  title?: string;
  description?: string;
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

  public get name(): string {
    return commands.SITE_CLASSIC_SET;
  }

  public get description(): string {
    return 'Change classic site settings';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.showDeprecationWarning(logger, commands.SITE_CLASSIC_SET, commands.SITE_SET);

    this.dots = '';

    spo
      .getTenantId(logger, this.debug)
      .then((_tenantId: string): Promise<string> => {
        this.tenantId = _tenantId;

        return spo.getSpoAdminUrl(logger, this.debug);
      })
      .then((_spoAdminUrl: string): Promise<FormDigestInfo> => {
        this.spoAdminUrl = _spoAdminUrl;

        return spo.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);
      })
      .then((res: FormDigestInfo): Promise<string> => {
        this.context = res;
        if (this.verbose) {
          logger.logToStderr(`Setting basic properties ${args.options.url}...`);
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
          updates.push(`<SetProperty Id="${++i}" ObjectPathId="5" Name="Title"><Parameter Type="String">${formatting.escapeXml(args.options.title)}</Parameter></SetProperty>`);
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
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${updates.join('')}<ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|${(this.tenantId as string).substr(pos, (this.tenantId as string).indexOf('&') - pos)}&#xA;SiteProperties&#xA;${encodeURIComponent(args.options.url)}" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`
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
            const isComplete: boolean = operation.IsComplete;
            if (!args.options.wait || isComplete) {
              resolve();
              return;
            }

            setTimeout(() => {
              spo.waitUntilFinished({
                operationId: JSON.stringify(operation._ObjectIdentity_),
                siteUrl: this.spoAdminUrl as string,
                resolve,
                reject,
                logger,
                currentContext: this.context as FormDigestInfo,
                dots: this.dots,
                debug: this.debug,
                verbose: this.verbose
              });
            }, operation.PollingInterval);
          }
        });
      })
      .then(_ => spo.ensureFormDigest(this.spoAdminUrl as string, logger, this.context, this.debug))
      .then((res: FormDigestInfo): Promise<void> => {
        this.context = res;
        return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
          if (!args.options.owners) {
            resolve();
            return;
          }

          Promise.all(args.options.owners.split(',').map(o => {
            return this.setAdmin(logger, args.options.url, o.trim());
          }))
            .then((): void => {
              resolve();
            }, (err: any): void => {
              reject(err);
            });
        });
      })
      .then(_ => spo.ensureFormDigest(this.spoAdminUrl as string, logger, this.context, this.debug))
      .then((res: FormDigestInfo): Promise<void> => {
        this.context = res;

        if (!args.options.description) {
          return Promise.resolve(undefined as any);
        }

        const requestOptions: any = {
          url: `${args.options.url}/_api/web`,
          headers: {
            'IF-MATCH': '*',
            'Accept': 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata',
            'X-RequestDigest': this.context.FormDigestValue,
            'X-HTTP-Method': 'MERGE'
          },
          data: {
            Description: args.options.description
          },
          json: true
        };

        return request.post(requestOptions);
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
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="7" ObjectPathId="5" Name="LockState"><Parameter Type="String">${formatting.escapeXml(args.options.lockState)}</Parameter></SetProperty><ObjectPath Id="9" ObjectPathId="8" /><ObjectIdentityQuery Id="10" ObjectPathId="5" /><Query Id="11" ObjectPathId="8"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="5" ParentId="3" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Id="8" ParentId="5" Name="Update" /><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
            const isComplete: boolean = operation.IsComplete;
            if (!args.options.wait || isComplete) {
              resolve();
              return;
            }

            setTimeout(() => {
              spo.waitUntilFinished({
                operationId: JSON.stringify(operation._ObjectIdentity_),
                siteUrl: this.spoAdminUrl as string,
                resolve,
                reject,
                logger,
                currentContext: this.context as FormDigestInfo,
                dots: this.dots,
                debug: this.debug,
                verbose: this.verbose
              });
            }, operation.PollingInterval);
          }
        });
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  private setAdmin(logger: Logger, siteUrl: string, principal: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      spo
        .ensureFormDigest(this.spoAdminUrl as string, logger, this.context, this.debug)
        .then((res: FormDigestInfo): Promise<string> => {
          this.context = res;
          const requestOptions: any = {
            url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': this.context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">${formatting.escapeXml(siteUrl)}</Parameter><Parameter Type="String">${formatting.escapeXml(principal)}</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
        option: '-u, --url <url>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--sharing [sharing]',
        autocomplete: ['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly']
      },
      {
        option: '--resourceQuota [resourceQuota]'
      },
      {
        option: '--resourceQuotaWarningLevel [resourceQuotaWarningLevel]'
      },
      {
        option: '--storageQuota [storageQuota]'
      },
      {
        option: '--storageQuotaWarningLevel [storageQuotaWarningLevel]'
      },
      {
        option: '--allowSelfServiceUpgrade [allowSelfServiceUpgrade]'
      },
      {
        option: '--owners [owners]'
      },
      {
        option: '--lockState [lockState]',
        autocomplete: ['Unlock', 'NoAdditions', 'ReadOnly', 'NoAccess']
      },
      {
        option: '--noScriptSite [noScriptSite]'
      },
      {
        option: '--wait'
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
  }
}

module.exports = new SpoSiteClassicSetCommand();