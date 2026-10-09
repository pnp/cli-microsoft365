import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { ListItemInstance } from '../listitem/ListItemInstance.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  title: z.string().optional().alias('t'),
  id: z.string().refine(val => !isNaN(Number(val)), { message: 'id is not a number' }).optional().alias('i'),
  clientSideComponentId: z.string().refine(val => validation.isValidGuid(val), { message: 'clientSideComponentId is not a valid GUID' }).optional().alias('c'),
  tenantWideExtensionComponentProperties: z.boolean().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTenantApplicationCustomizerGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_GET;
  }

  public get description(): string {
    return 'Gets an application customizer that is installed tenant wide';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.title, opts.id, opts.clientSideComponentId].filter(v => v !== undefined).length === 1, {
      error: `Specify exactly one of the following options: 'title', 'id', or 'clientSideComponentId'.`
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

      if (!appCatalogUrl) {
        throw 'No app catalog URL found';
      }

      let filter: string;
      if (args.options.title) {
        filter = `Title eq '${args.options.title}'`;
      }
      else if (args.options.id) {
        filter = `Id eq '${args.options.id}'`;
      }
      else {
        filter = `TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
      }

      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
      const listItemInstances = await odata.getAllItems<ListItemInstance>(`${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and ${filter}`);

      if (listItemInstances) {
        if (listItemInstances.length === 0) {
          throw 'The specified application customizer was not found';
        }

        listItemInstances.forEach(v => delete (v as any)['ID']);

        let listItemInstance: ListItemInstance;
        if (listItemInstances.length > 1) {
          const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances);
          listItemInstance = await cli.handleMultipleResultsFound<ListItemInstance>(`Multiple application customizers with ${args.options.title || args.options.clientSideComponentId} were found.`, resultAsKeyValuePair);
        }
        else {
          listItemInstance = listItemInstances[0];
        }

        if (!args.options.tenantWideExtensionComponentProperties) {
          await logger.log(listItemInstance);
        }
        else {
          const properties = formatting.tryParseJson((listItemInstance as any).TenantWideExtensionComponentProperties);
          await logger.log(properties);
        }
      }
      else {
        throw 'The specified application customizer was not found';
      }
    }
    catch (err: any) {
      return this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantApplicationCustomizerGetCommand();