import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListItemInstance } from '../listitem/ListItemInstance.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  title: z.string().optional().alias('t'),
  id: z.string().refine(val => !isNaN(Number(val)), { message: 'id is not a valid list item ID' }).optional().alias('i'),
  clientSideComponentId: z.string().refine(val => validation.isValidGuid(val), { message: 'clientSideComponentId is not a valid GUID' }).optional().alias('c'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTenantApplicationCustomizerRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPLICATIONCUSTOMIZER_REMOVE;
  }

  public get description(): string {
    return 'Removes an application customizer that is installed tenant wide.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.title, opts.id, opts.clientSideComponentId].filter(v => v !== undefined).length === 1, {
      error: `Specify exactly one of the following options: 'title', 'id', or 'clientSideComponentId'.`,
      params: {
        customCode: 'optionSet',
        options: ['title', 'id', 'clientSideComponentId']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.force) {
        return await this.removeTenantApplicationCustomizer(logger, args);
      }

      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the tenant applicationcustomizer ${args.options.id || args.options.title || args.options.clientSideComponentId}?` });

      if (result) {
        await this.removeTenantApplicationCustomizer(logger, args);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  public async getTenantApplicationCustomizerId(logger: Logger, args: CommandArgs, requestUrl: string): Promise<number> {
    if (this.verbose) {
      await logger.logToStderr(`Getting the tenant application customizer ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
    }

    const filter: string[] = [`TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer'`];
    if (args.options.title) {
      filter.push(`Title eq '${args.options.title}'`);
    }
    else if (args.options.id) {
      filter.push(`Id eq '${args.options.id}'`);
    }
    else if (args.options.clientSideComponentId) {
      filter.push(`TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`);
    }

    const listItemInstances: ListItemInstance[] = await odata.getAllItems(`${requestUrl}/items?$filter=${filter.join(' and ')}&$select=Id`);

    if (listItemInstances.length === 0) {
      throw 'The specified application customizer was not found';
    }

    if (listItemInstances.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances);
      listItemInstances[0] = await cli.handleMultipleResultsFound<ListItemInstance>(`Multiple application customizers with ${args.options.title || args.options.clientSideComponentId} were found.`, resultAsKeyValuePair);
    }

    return listItemInstances[0].Id;
  }

  private async removeTenantApplicationCustomizer(logger: Logger, args: CommandArgs): Promise<void> {
    const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

    if (!appCatalogUrl) {
      throw 'No app catalog URL found';
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
    const requestUrl = `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    const id = await this.getTenantApplicationCustomizerId(logger, args, requestUrl);

    if (this.verbose) {
      await logger.logToStderr(`Removing tenant application customizer ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}/items(${id})`,
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*',
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }
}

export default new SpoTenantApplicationCustomizerRemoveCommand();