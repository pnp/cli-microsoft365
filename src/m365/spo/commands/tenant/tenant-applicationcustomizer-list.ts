import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListItemInstanceCollection } from '../listitem/ListItemInstanceCollection.js';

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

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');

    try {
      const response = await odata.getAllItems<ListItemInstanceCollection>(`${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer'`);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantApplicationCustomizerListCommand();