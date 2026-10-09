import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError, globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListItemInstance } from '../listitem/ListItemInstance';
import { ListItemInstanceCollection } from '../listitem/ListItemInstanceCollection.js';

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

class SpoTenantCommandSetGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_COMMANDSET_GET;
  }

  public get description(): string {
    return 'Gets a ListView Command Set that is installed tenant wide';
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
    const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
    if (!appCatalogUrl) {
      throw new CommandError('No app catalog URL found');
    }

    let filter: string = `startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet')`;

    if (args.options.title) {
      filter += ` and Title eq '${args.options.title}'`;
    }
    else if (args.options.id) {
      filter += ` and Id eq ${args.options.id}`;
    }
    else if (args.options.clientSideComponentId) {
      filter += ` and TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
    const reqOptions: CliRequestOptions = {
      url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=${filter}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const listItemInstances = await request.get<ListItemInstanceCollection>(reqOptions);

      if (listItemInstances?.value.length > 0) {
        listItemInstances.value.forEach(v => delete v['ID']);

        let listItemInstance: ListItemInstance;
        if (listItemInstances.value.length > 1) {
          const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances.value);
          listItemInstance = await cli.handleMultipleResultsFound<ListItemInstance>(`Multiple ListView Command Sets with ${args.options.title || args.options.clientSideComponentId} were found.`, resultAsKeyValuePair);
        }
        else {
          listItemInstance = listItemInstances.value[0];
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
        throw 'The specified ListView Command Set was not found';
      }
    }
    catch (err: any) {
      return this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantCommandSetGetCommand();