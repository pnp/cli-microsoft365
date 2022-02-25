import { Group } from '@microsoft/microsoft-graph-types';
import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
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
          .then((groupId: string) => {
            if (groupId === '00000000-0000-0000-0000-000000000000') {
              if (this.debug) {
                logger.logToStderr('Site is not groupified. Going ahead with the conventional site deletion options');
              }

              return this.deleteSiteWithoutGroup(logger, args);
            }
            else {
              if (this.debug) {
                logger.logToStderr(`Site attached to group ${groupId}. Initiating group delete operation via Graph API`);
              }

              return this
                .getSiteGroup(groupId)
                .then((group) => {
                  if (args.options.skipRecycleBin || args.options.wait) {
                    logger.logToStderr(chalk.yellow(`Entered site is a groupified site. Hence, the parameters 'skipRecycleBin' and 'wait' will not be applicable.`));
                  }

                  return this.deleteGroupifiedSite(group.id, logger);
                })
                .catch((err: any) => {
                  if (err.response.status === 404) {
                    if (this.verbose) {
                      logger.logToStderr(`Site group doesn't exist. Searching in the Microsoft 365 deleted groups.`);
                    }

                    return this
                      .isSiteGroupDeleted(groupId)
                      .then((deletedGroups: { value: { id: string }[] }): Promise<void> => {
                        if (deletedGroups.value.length === 0) {
                          if (this.verbose) {
                            logger.logToStderr("Site group doesn't exist anymore. Deleting the site.");
                          }

                          if (args.options.wait) {
                            logger.logToStderr(chalk.yellow(`Entered site is a groupified site. Hence, the parameter 'wait' will not be applicable.`));
                          }

                          return Promise.resolve();
                        }
                        else {
                          return Promise.reject(`Site group still exists in the deleted groups. The site won't be removed.`);
                        }
                      })
                      .then(_ => this.deleteOrphanedSite(logger, args.options.url))
                      .catch((err) => Promise.reject(err));
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

  private isSiteGroupDeleted(groupId: string): Promise<{ value: { id: string }[] }> {
    const requestOptions: any = {
      url: `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$select=id&$filter=groupTypes/any(c:c+eq+'Unified') and startswith(id, '${groupId}')`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<{ value: { id: string }[] }>(requestOptions);
  }

  private deleteOrphanedSite(logger: Logger, url: string): Promise<void> {
    return spo
      .getSpoAdminUrl(logger, this.debug)
      .then((spoAdminUrl: string): Promise<void> => {
        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/GroupSiteManager/Delete?siteUrl='${url}'`,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      });
  }

  private deleteSiteWithoutGroup(logger: Logger, args: CommandArgs): Promise<void> {
    return spo
      .getSpoAdminUrl(logger, this.debug)
      .then((spoAdminUrl: string): Promise<FormDigestInfo> => {
        this.spoAdminUrl = spoAdminUrl;

        return spo.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);
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
      spo
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
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54"/><ObjectPath Id="57" ObjectPathId="56"/><Query Id="58" ObjectPathId="54"><Query SelectAllProperties="true"><Properties/></Query></Query><Query Id="59" ObjectPathId="56"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true"/><Property Name="PollingInterval" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="54" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="56" ParentId="54" Name="RemoveSite"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method></ObjectPaths></Request>`
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
    });
  }

  private deleteSiteFromTheRecycleBin(url: string, wait: boolean, logger: Logger): Promise<void> {
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
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
    });
  }

  private getSiteGroupId(url: string, logger: Logger): Promise<string> {
    return spo
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<FormDigestInfo> => {
        this.spoAdminUrl = _spoAdminUrl;
        return spo.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);
      })
      .then((res: FormDigestInfo): Promise<string> => {
        this.context = res;
        if (this.verbose) {
          logger.logToStderr(`Retrieving the group Id of the site ${url}`);
        }

        const requestOptions: any = {
          url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': this.context.FormDigestValue
          },
          data: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="4" ObjectPathId="3"/><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="GroupId" ScalarProperty="true"/></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/><Method Id="3" ParentId="1" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method></ObjectPaths></Request>`
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

  private deleteGroupifiedSite(groupId: string | undefined, logger: Logger): Promise<void> {
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
    return validation.isValidSharePointUrl(args.options.url);
  }
}

module.exports = new SpoSiteRemoveCommand();