import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { setTimeout } from 'timers/promises';
import { urlUtil } from '../../../../utils/urlUtil.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  skipRecycleBin?: boolean;
  fromRecycleBin?: boolean;
  wait?: boolean;
  force?: boolean;
}

interface SiteDetails {
  GroupId: string,
  TimeDeleted?: string,
  SiteId: string
}

class SpoSiteRemoveCommand extends SpoCommand {
  private spoAdminUrl?: string;
  private pollingInterval = 5000;

  public get name(): string {
    return commands.SITE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        skipRecycleBin: !!args.options.skipRecycleBin,
        fromRecycleBin: !!args.options.fromRecycleBin,
        wait: !!args.options.wait,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        const uri = new URL(args.options.url);
        const rootUrl = `${uri.protocol}//${uri.hostname}`;
        if (rootUrl.toLowerCase() === urlUtil.removeTrailingSlashes(args.options.url.toLowerCase())) {
          return `The root site cannot be deleted.`;
        }

        if (args.options.fromRecycleBin && args.options.skipRecycleBin) {
          return 'Specify either fromRecycleBin or skipRecycleBin, but not both.';
        }

        return true;
      });
  }

  #initTypes(): void {
    this.types.string.push('url');
    this.types.boolean.push('skipRecycleBin', 'fromRecycleBin', 'wait', 'force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.wait) {
      await this.warn(logger, `Option 'wait' is deprecated and will be removed in the next major release.`);
    }

    if (args.options.force) {
      await this.removeSite(logger, args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the site '${args.options.url}'?` });

      if (result) {
        await this.removeSite(logger, args.options);
      }
    }
  }

  private async removeSite(logger: Logger, options: Options): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Removing site '${options.url}'...`);
      }

      this.spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      const siteUrl = urlUtil.removeTrailingSlashes(options.url);
      const siteDetails: SiteDetails = await this.getSiteDetails(logger, siteUrl);
      const isGroupSite = siteDetails.GroupId && siteDetails.GroupId !== '00000000-0000-0000-0000-000000000000';

      if (options.fromRecycleBin) {
        if (!siteDetails.TimeDeleted) {
          throw `Site is currently not in the recycle bin. Remove --fromRecycleBin if you want to remove it as active site.`;
        }

        if (isGroupSite) {
          if (this.verbose) {
            await logger.logToStderr(`Checking if group '${siteDetails.GroupId}' is already permanently deleted from recycle bin.`);
          }

          const isGroupInRecycleBin = await this.isGroupInEntraRecycleBin(logger, siteDetails.GroupId);
          if (isGroupInRecycleBin) {
            await this.removeGroupFromEntraRecycleBin(logger, siteDetails.GroupId);
          }
        }
        await this.deleteSiteFromSharePointRecycleBin(logger, siteUrl);
      }
      else {
        if (siteDetails.TimeDeleted) {
          throw `Site is already in the recycle bin. Use --fromRecycleBin to permanently delete it.`;
        }

        if (isGroupSite) {
          await this.deleteGroupifiedSite(logger, siteUrl);
          if (options.skipRecycleBin) {
            let isGroupInRecycleBin = await this.isGroupInEntraRecycleBin(logger, siteDetails.GroupId);
            let amountOfPolls = 0;

            while (!isGroupInRecycleBin && amountOfPolls < 20) {
              await setTimeout(this.pollingInterval);
              isGroupInRecycleBin = await this.isGroupInEntraRecycleBin(logger, siteDetails.GroupId);
              amountOfPolls++;
            }

            if (isGroupInRecycleBin) {
              await this.removeGroupFromEntraRecycleBin(logger, siteDetails.GroupId);
            }
          }
        }
        else {
          await this.deleteNonGroupSite(logger, siteUrl);
        }

        if (options.skipRecycleBin) {
          await this.deleteSiteFromSharePointRecycleBin(logger, siteUrl);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeGroupFromEntraRecycleBin(logger: Logger, groupId: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Permanently deleting group '${groupId}'.`);
    }

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${groupId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.delete(requestOptions);
  }

  private async isGroupInEntraRecycleBin(logger: Logger, groupId: string): Promise<boolean> {
    if (this.verbose) {
      await logger.logToStderr(`Checking if group '${groupId}' is in the Microsoft Entra recycle bin.`);
    }

    const requestOptions: CliRequestOptions = {
      url: `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${groupId}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      await request.get(requestOptions);
      return true;
    }
    catch (err: any) {
      if (err.response?.status === 404) {
        return false;
      }
      throw err;
    }
  }

  private async deleteNonGroupSite(logger: Logger, siteUrl: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Deleting site.`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.spoAdminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveSite`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: {
        siteUrl: siteUrl
      },
      responseType: 'json'
    };
    return request.post(requestOptions);
  }

  private async deleteSiteFromSharePointRecycleBin(logger: Logger, url: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Permanently deleting site from the recycle bin.`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.spoAdminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveDeletedSite`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json'
      },
      data: {
        siteUrl: url
      },
      responseType: 'json'
    };
    return request.post(requestOptions);
  }

  private async getSiteDetails(logger: Logger, url: string): Promise<SiteDetails> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving site info.`);
    }

    const sites = await odata.getAllItems<SiteDetails>(`${this.spoAdminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/items?$filter=SiteUrl eq '${formatting.encodeQueryParameter(url)}'&$select=GroupId,TimeDeleted,SiteId`);

    if (sites.length === 0) {
      throw `Site not found in the tenant.`;
    }
    return sites[0];
  }

  private async deleteGroupifiedSite(logger: Logger, siteUrl: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing groupified site.`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.spoAdminUrl}/_api/GroupSiteManager/Delete?siteUrl='${formatting.encodeQueryParameter(siteUrl)}'`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }
}

export default new SpoSiteRemoveCommand();