import { Logger } from '../../../../cli/Logger';
import { CommandError } from '../../../../Command';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListItemInstanceCollection } from '../listitem/ListItemInstanceCollection';

class SpoTenantCommandSetListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_COMMANDSET_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of ListView Command Sets that are installed tenant-wide';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'TenantWideExtensionComponentId', 'TenantWideExtensionListTemplate'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

    if (!appCatalogUrl) {
      throw new CommandError('No app catalog URL found');
    }

    if (this.verbose) {
      logger.logToStderr('Retrieving a list of ListView Command Sets that are installed tenant-wide');
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');

    try {
      const response = await odata.getAllItems<ListItemInstanceCollection>(`${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet')`);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoTenantCommandSetListCommand();