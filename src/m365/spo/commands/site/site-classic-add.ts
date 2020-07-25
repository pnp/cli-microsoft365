import config from '../../../../config';
import commands from '../../commands';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CommandOption, CommandValidate } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo } from '../../spo';
import { SiteProperties } from './SiteProperties';
import { DeletedSiteProperties } from './DeletedSiteProperties';
import { SpoOperation } from './SpoOperation';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  title: string;
  owner: string;
  timeZone: string | number;
  lcid?: string | number;
  webTemplate?: string;
  resourceQuota?: string | number;
  resourceQuotaWarningLevel?: string | number;
  storageQuota?: string | number;
  storageQuotaWarningLevel?: string | number;
  removeDeletedSite: boolean;
  wait: boolean;
}

class SpoSiteClassicAddCommand extends SpoCommand {
  private context?: FormDigestInfo;
  private spoAdminUrl?: string;
  private dots?: string;
  private timeout?: NodeJS.Timer;

  public get name(): string {
    return commands.SITE_CLASSIC_ADD;
  }

  public get description(): string {
    return 'Creates new classic site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.lcid = typeof args.options.lcid !== 'undefined';
    telemetryProps.webTemplate = typeof args.options.webTemplate !== 'undefined';
    telemetryProps.resourceQuota = typeof args.options.resourceQuota !== 'undefined';
    telemetryProps.resourceQuotaWarningLevel = typeof args.options.resourceQuotaWarningLevel !== 'undefined';
    telemetryProps.storageQuota = typeof args.options.storageQuota !== 'undefined';
    telemetryProps.storageQuotaWarningLevel = typeof args.options.storageQuotaWarningLevel !== 'undefined';
    telemetryProps.removeDeletedSite = args.options.removeDeletedSite;
    telemetryProps.wait = args.options.wait;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this.dots = '';
    this.showDeprecationWarning(cmd, commands.SITE_CLASSIC_ADD, commands.SITE_ADD); 

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<FormDigestInfo> => {
        this.spoAdminUrl = _spoAdminUrl;

        return this.ensureFormDigest(this.spoAdminUrl, cmd, this.context, this.debug);
      })
      .then((res: FormDigestInfo): Promise<boolean> => {
        this.context = res;

        if (args.options.removeDeletedSite) {
          return this.siteExistsInTheRecycleBin(args.options.url, cmd);
        }
        else {
          // assume site doesn't exist
          return Promise.resolve(false);
        }
      })
      .then((exists: boolean): Promise<void> => {
        if (exists) {
          if (this.verbose) {
            cmd.log('Site exists in the recycle bin');
          }

          return this.deleteSiteFromTheRecycleBin(args.options.url, args.options.wait, cmd);
        }
        else {
          if (this.verbose) {
            cmd.log('Site not found');
          }

          return Promise.resolve();
        }
      })
      .then((): Promise<FormDigestInfo> => {
        return this.ensureFormDigest(this.spoAdminUrl as string, cmd, this.context, this.debug);
      })
      .then((res: FormDigestInfo): Promise<string> => {
        this.context = res;

        if (this.verbose) {
          cmd.log(`Creating site collection ${args.options.url}...`);
        }

        const lcid: number = typeof args.options.lcid === 'number' ? args.options.lcid : 1033;
        const storageQuota: number = typeof args.options.storageQuota === 'number' ? args.options.storageQuota : 100;
        const storageQuotaWarningLevel: number = typeof args.options.storageQuotaWarningLevel === 'number' ? args.options.storageQuotaWarningLevel : 100;
        const resourceQuota: number = typeof args.options.resourceQuota === 'number' ? args.options.resourceQuota : 0;
        const resourceQuotaWarningLevel: number = typeof args.options.resourceQuotaWarningLevel === 'number' ? args.options.resourceQuotaWarningLevel : 0;
        const webTemplate: string = args.options.webTemplate || 'STS#0';

        const requestOptions: any = {
          url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': this.context.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">${lcid}</Property><Property Name="Owner" Type="String">${Utils.escapeXml(args.options.owner)}</Property><Property Name="StorageMaximumLevel" Type="Int64">${storageQuota}</Property><Property Name="StorageWarningLevel" Type="Int64">${storageQuotaWarningLevel}</Property><Property Name="Template" Type="String">${Utils.escapeXml(webTemplate)}</Property><Property Name="TimeZoneId" Type="Int32">${args.options.timeZone}</Property><Property Name="Title" Type="String">${Utils.escapeXml(args.options.title)}</Property><Property Name="Url" Type="String">${Utils.escapeXml(args.options.url)}</Property><Property Name="UserCodeMaximumLevel" Type="Double">${resourceQuota}</Property><Property Name="UserCodeWarningLevel" Type="Double">${resourceQuotaWarningLevel}</Property></Parameter></Parameters></Method></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): Promise<void> => {
        return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
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

  private siteExistsInTheRecycleBin(url: string, cmd: CommandInstance): Promise<boolean> {
    return new Promise<boolean>((resolve: (exists: boolean) => void, reject: (error: any) => void): void => {
      this
        .ensureFormDigest(this.spoAdminUrl as string, cmd, this.context, this.debug)
        .then((res: FormDigestInfo): Promise<string> => {
          this.context = res;

          if (this.verbose) {
            cmd.log(`Checking if the site ${url} exists...`);
          }

          const requestOptions: any = {
            url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': this.context.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(url)}</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`
          };

          return request.post(requestOptions);
        })
        .then((res: string): Promise<boolean> => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            if (response.ErrorInfo.ErrorTypeName === 'Microsoft.Online.SharePoint.Common.SpoNoSiteException') {
              return Promise.resolve(false);
            }
            else {
              return Promise.reject(response.ErrorInfo.ErrorMessage);
            }
          }
          else {
            const site: SiteProperties = json[json.length - 1];
            if (site.Status === 'Recycled') {
              return Promise.reject(true);
            }
            else {
              return Promise.resolve(false);
            }
          }
        })
        .then((exists: boolean): Promise<string> => {
          if (this.verbose) {
            cmd.log(`Site doesn't exist. Checking if the site ${url} exists in the recycle bin...`);
          }

          const requestOptions: any = {
            url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': (this.context as FormDigestInfo).FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            if (response.ErrorInfo.ErrorTypeName === 'Microsoft.SharePoint.Client.UnknownError') {
              resolve(false);
            }
            else {
              reject(response.ErrorInfo.ErrorMessage);
            }
          }
          else {
            const site: DeletedSiteProperties = json[json.length - 1];
            if (site.Status === 'Recycled') {
              resolve(true);
            }
            else {
              resolve(false);
            }
          }
        }, (error: any): void => {
          if (typeof error === 'boolean') {
            resolve(error);
          }
          else {
            reject(error);
          }
        });
    });
  }

  private deleteSiteFromTheRecycleBin(url: string, wait: boolean, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .ensureFormDigest(this.spoAdminUrl as string, cmd, this.context, this.debug)
        .then((res: FormDigestInfo): Promise<string> => {
          this.context = res;

          if (this.verbose) {
            cmd.log(`Deleting site ${url} from the recycle bin...`);
          }

          const requestOptions: any = {
            url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': this.context.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${Utils.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
            const operation: SpoOperation = json[json.length - 1];
            let isComplete: boolean = operation.IsComplete;
            if (!wait || isComplete) {
              resolve();
              return;
            }

            setTimeout(() => {
              this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), this.spoAdminUrl as string, resolve, reject, cmd, this.context as FormDigestInfo, this.dots, this.timeout);
            }, operation.PollingInterval);
          }
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
        option: '-t, --title <title>',
        description: 'The site title'
      },
      {
        option: '--owner <owner>',
        description: 'The account name of the site owner'
      },
      {
        option: '-z, --timeZone <timeZone>',
        description: 'Integer representing time zone to use for the site'
      },
      {
        option: '-l, --lcid [lcid]',
        description: 'Integer representing time zone to use for the site'
      },
      {
        option: '-w, --webTemplate [webTemplate]',
        description: 'Template to use for creating the site. Default STS#0'
      },
      {
        option: '--resourceQuota [resourceQuota]',
        description: 'The quota for this site collection in Sandboxed Solutions units. Default 0'
      },
      {
        option: '--resourceQuotaWarningLevel [resourceQuotaWarningLevel]',
        description: 'The warning level for the resource quota. Default 0'
      },
      {
        option: '--storageQuota [storageQuota]',
        description: 'The storage quota for this site collection in megabytes. Default 100'
      },
      {
        option: '--storageQuotaWarningLevel [storageQuotaWarningLevel]',
        description: 'The warning level for the storage quota in megabytes. Default 100'
      },
      {
        option: '--removeDeletedSite',
        description: 'Set, to remove existing deleted site with the same URL from the Recycle Bin'
      },
      {
        option: '--wait',
        description: 'Wait for the site to be provisioned before completing the command'
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

      if (typeof args.options.timeZone !== 'number') {
        return `${args.options.timeZone} is not a number`;
      }

      if (args.options.lcid &&
        typeof args.options.lcid !== 'number') {
        return `${args.options.lcid} is not a number`;
      }

      if (args.options.resourceQuota &&
        typeof args.options.resourceQuota !== 'number') {
        return `${args.options.resourceQuota} is not a number`;
      }

      if (args.options.resourceQuotaWarningLevel &&
        typeof args.options.resourceQuotaWarningLevel !== 'number') {
        return `${args.options.resourceQuotaWarningLevel} is not a number`;
      }

      if (args.options.resourceQuotaWarningLevel &&
        !args.options.resourceQuota) {
        return `You cannot specify resourceQuotaWarningLevel without specifying resourceQuota`;
      }

      if ((<number>args.options.resourceQuotaWarningLevel) > (<number>args.options.resourceQuota)) {
        return `resourceQuotaWarningLevel cannot exceed resourceQuota`;
      }

      if (args.options.storageQuota &&
        typeof args.options.storageQuota !== 'number') {
        return `${args.options.storageQuota} is not a number`;
      }

      if (args.options.storageQuotaWarningLevel &&
        typeof args.options.storageQuotaWarningLevel !== 'number') {
        return `${args.options.storageQuotaWarningLevel} is not a number`;
      }

      if (args.options.storageQuotaWarningLevel &&
        !args.options.storageQuota) {
        return `You cannot specify storageQuotaWarningLevel without specifying storageQuota`;
      }

      if ((<number>args.options.storageQuotaWarningLevel) > (<number>args.options.storageQuota)) {
        return `storageQuotaWarningLevel cannot exceed storageQuota`;
      }

      return true;
    };
  }
}

module.exports = new SpoSiteClassicAddCommand();