import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { ListItemListOptions, spoListItem } from '../../../../utils/spoListItem.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

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
      await logger.logToStderr('Retrieving a list of ListView Command Sets that are installed tenant-wide');
    }

    try {
      const options: ListItemListOptions = {
        webUrl: appCatalogUrl,
        listUrl: '/Lists/TenantWideExtensions',
        filter: `startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet')`
      };

      const listItems = await spoListItem.getListItems(options, logger, this.verbose);

      await logger.log(listItems);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantCommandSetListCommand();