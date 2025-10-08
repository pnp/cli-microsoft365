import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { ListItemListOptions, spoListItem } from '../../../../utils/spoListItem.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

class SpoTenantApplicationCustomizerListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of application customizers that are installed tenant-wide.';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'TenantWideExtensionComponentId', 'TenantWideExtensionWebTemplate'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

    if (!appCatalogUrl) {
      throw new CommandError('No app catalog URL found');
    }

    try {
      const options: ListItemListOptions = {
        webUrl: appCatalogUrl,
        listUrl: '/Lists/TenantWideExtensions',
        filter: `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer'`
      };

      const listItems = await spoListItem.getListItems(options, logger, this.verbose);

      await logger.log(listItems);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantApplicationCustomizerListCommand();