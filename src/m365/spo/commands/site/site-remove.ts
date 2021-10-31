import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { Group } from '../../../aad/commands/o365group/Group';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo } from '../../spo';
import { SpoOperation } from './SpoOperation';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  skipRecycleBin?: boolean;
  fromRecycleBin?: boolean;
  wait: boolean;
  confirm?: boolean;
}

class SpoSiteRemoveCommand extends SpoCommand {
  private context?: FormDigestInfo;
  private spoAdminUrl?: string;
  private dots?: string;

  public get name(): string {
    return commands.SITE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.skipRecycleBin = (!(!args.options.skipRecycleBin)).toString();
    telemetryProps.fromRecycleBin = (!(!args.options.fromRecycleBin)).toString();
    telemetryProps.wait = args.options.wait;
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeSite = (): void => {
      this.dots = '';

      if (args.options.fromRecycleBin) {
        this
          .deleteSiteWithoutGroup(logger, args)
          .then(_ => cb(), (err: any): void => this.handleRejectedPromise(err, logger, cb));
      }
      else {
        this
          .getSiteGroupId(args.options.url, logger)
          .then((_groupId: string) => {
            if (_groupId === '00000000-0000-0000-0000-000000000000') {
              if (this.debug) {
                logger.logToStderr('Site is not groupified. Going ahead with the conventional site deletion options');
              }

              return this.deleteSiteWithoutGroup(logger, args);
            }
            else {
              if (this.debug) {
                logger.logToStderr(`Site attached to group ${_groupId}. Initiating group delete operation via Graph API`);
              }

              return this.getSiteGroup(_groupId)
                .then((grp) => {
                  if (args.options.skipRecycleBin || args.options.wait) {
                    logger.log(chalk.yellow(`Entered site is a groupified site. Hence, the parameters 'skipRecycleBin' and 'wait' will not be applicable.`));
                  }

                  return this.deleteGroupifiedSite(grp.id, logger);
                })
                .catch((err: any) => {
                  if (err.response.status === 404) {
                    if (this.verbose) {
                      logger.log(`Site group doesn't exist. Searching in the Microsoft 365 deleted groups.`);
                    }
        
                    return this.isSiteGroupDeleted(_groupId)
                      .then((deletedGrps: any): Promise<void> => {
                        if (deletedGrps.value.length === 0) {
                          if (this.verbose) {
                            logger.log("Site group doesn't exist anymore. Deleting the site.");
                          }

                          if (args.options.wait) {
                            logger.log(chalk.yellow(`Entered site is a groupified site. Hence, the parameter 'wait' will not be applicable.`));
                          }

                          return Promise.resolve();
                        }
                        else {
                          return Promise.reject(`Site group still exists in the deleted groups. The site won't be removed.`);
                        }
                      })
                      .then(() => {
                        return this.deleteOrphanedSite(logger, args.options.url);
                      })
                      .catch((err) => {
                        return Promise.reject(err);
                      });
                  }
                  else {
                    return Promise.reject(err);
                  }
                });
            }
          })
          .then(_ => cb(), (err: any): void => this.handleRejectedPromise(err, logger, cb));
      }
    };

    if (args.options.confirm) {
      removeSite();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the site ${args.options.url}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeSite();
        }
      });
    }
  }

  private getSiteGroup(groupId: string): Promise<Group> {
    const requestOptions: any = {
      url: `https://graph.microsoft.com/v1.0/groups/${groupId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<Group>(requestOptions);
  }

  private isSiteGroupDeleted(groupId: string): Promise<any> {
    const requestOptions: any = {
      url: `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(id, '${groupId}')`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<any>(requestOptions);
  }

  private deleteOrphanedSite(logger: Logger, url: string): Promise<void> {
    return this
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<void> => {
        const requestOptions: any = {
          url: `${_spoAdminUrl}/_api/GroupSiteManager/Delete?siteUrl='${url}'`,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata',
            responseType: 'json'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .catch((err) => {
        return Promise.reject(err);
      });
  }

  private deleteSiteWithoutGroup(logger: Logger, args: CommandArgs): Promise<void> {
    return this
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<FormDigestInfo> => {
        this.spoAdminUrl = _spoAdminUrl;

        return this.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);
      })
      .then((res: FormDigestInfo): Promise<void> => {
        this.context = res;

        if (args.options.fromRecycleBin) {
          if (this.verbose) {
            logger.logToStderr(`Deleting site from recycle bin ${args.options.url}...`);
          }

          return this.deleteSiteFromTheRecycleBin(args.options.url, args.options.wait, logger);
        }
        else {
          return this.deleteSite(args.options.url, args.options.wait, logger);
        }
      })
      .then((): Promise<void> => {
        if (args.options.skipRecycleBin) {
          if (this.verbose) {
            logger.logToStderr(`Also deleting site from tenant recycle bin ${args.options.url}...`);
          }
          return this.deleteSiteFromTheRecycleBin(args.options.url, args.options.wait, logger);
        }
        else {
          return Promise.resolve();
        }
      });
  }

  private deleteSite(url: string, wait: boolean, logger: Logger): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .ensureFormDigest(this.spoAdminUrl as string, logger, this.context, this.debug)
        .then((res: FormDigestInfo): Promise<string> => {
          this.context = res;

          if (this.verbose) {
            logger.logToStderr(`Deleting site ${url}...`);
          }

          const requestOptions: any = {
            url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': this.context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">${Utils.escapeXml(url)}</Parameter></Parameters></Method></ObjectPaths></Request>`
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
            const isComplete: boolean = operation.IsComplete;
            if (!wait || isComplete) {
              resolve();
              return;
            }

            setTimeout(() => {
              this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), this.spoAdminUrl as string, resolve, reject, logger, this.context as FormDigestInfo, this.dots);
            }, operation.PollingInterval);
          }
        });
    });
  }

  private deleteSiteFromTheRecycleBin(url: string, wait: boolean, logger: Logger): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this
        .ensureFormDigest(this.spoAdminUrl as string, logger, this.context, this.debug)
        .then((res: FormDigestInfo): Promise<string> => {
          this.context = res;

          const requestOptions: any = {
            url: `${this.spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': this.context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${Utils.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
            const isComplete: boolean = operation.IsComplete;
            if (!wait || isComplete) {
              resolve();
              return;
            }

            setTimeout(() => {
              this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), this.spoAdminUrl as string, resolve, reject, logger, this.context as FormDigestInfo, this.dots);
            }, operation.PollingInterval);
          }
        });
    });
  }

  private getSiteGroupId(url: string, logger: Logger): Promise<string> {
    return this
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<FormDigestInfo> => {
        this.spoAdminUrl = _spoAdminUrl;
        return this.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);
      })
      .then((res: FormDigestInfo): Promise<string> => {
        this.context = res;
        if (this.verbose) {
          logger.logToStderr(`Retrieving the group Id of the site  ${url}`);
        }

        const requestOptions: any = {
          url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': this.context.FormDigestValue
          },
          data: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(url)}</Parameter></Parameters></Method></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): Promise<string> => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          return Promise.reject(response.ErrorInfo.ErrorMessage);
        }
        else {
          const groupId: string = json[json.length - 1].GroupId.replace('/Guid(', '').replace(')/', '');
          return Promise.resolve(groupId);
        }
      });
  }

  private deleteGroupifiedSite(groupId: string, logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing Microsoft 365 Group: ${groupId}...`);
    }

    const requestOptions: any = {
      url: `https://graph.microsoft.com/v1.0/groups/${groupId}`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      }
    };

    return request.delete(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '--skipRecycleBin'
      },
      {
        option: '--fromRecycleBin'
      },
      {
        option: '--wait'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.url);
  }
}

module.exports = new SpoSiteRemoveCommand();