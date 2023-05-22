import { Logger } from '../../../../cli/Logger';
import { CommandError } from '../../../../Command';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListItemInstanceCollection } from '../listitem/ListItemInstanceCollection';

class SpoTenantApplicationCustomizerListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_LIST;
  }

  public get description(): string {
    return 'Get a list of application customizers that are installed tenant wide';
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
    const reqOptions: CliRequestOptions = {
      url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer'`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<ListItemInstanceCollection>(reqOptions);
      if (res.value && res.value.length > 0) {
        logger.log(res.value);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No tenant wide installed application customizers found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoTenantApplicationCustomizerListCommand();