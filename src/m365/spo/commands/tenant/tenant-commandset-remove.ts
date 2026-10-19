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

class SpoTenantCommandSetRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_COMMANDSET_REMOVE;
  }

  public get description(): string {
    return 'Removes a ListView Command Set that is installed tenant wide.';
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
        return await this.removeTenantCommandSet(logger, args);
      }

      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the tenant commandset ${args.options.id || args.options.title || args.options.clientSideComponentId}?` });

      if (result) {
        await this.removeTenantCommandSet(logger, args);
      }
    }
    catch (err: any) {
      await logger.log(err);
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeTenantCommandSet(logger: Logger, args: CommandArgs): Promise<void> {
    const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

    if (!appCatalogUrl) {
      throw 'No app catalog URL found';
    }

    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
    const requestUrl = `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    const id = await this.getTenantCommandSetId(logger, args, requestUrl);

    if (this.verbose) {
      await logger.logToStderr(`Removing tenant command set ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
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

  public async getTenantCommandSetId(logger: Logger, args: CommandArgs, requestUrl: string): Promise<number> {
    if (this.verbose) {
      await logger.logToStderr(`Getting the tenant command set ${args.options.id || args.options.title || args.options.clientSideComponentId}`);
    }

    let filter: string;
    if (args.options.title) {
      filter = `Title eq '${args.options.title}'`;
    }
    else if (args.options.id) {
      filter = `Id eq ${args.options.id}`;
    }
    else {
      filter = `TenantWideExtensionComponentId eq '${args.options.clientSideComponentId}'`;
    }

    const listItemInstances: ListItemInstance[] = await odata.getAllItems<ListItemInstance>(`${requestUrl}/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and ${filter}`);

    if (listItemInstances.length === 0) {
      throw 'The specified command set was not found';
    }

    if (listItemInstances.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', listItemInstances);
      const result = await cli.handleMultipleResultsFound<ListItemInstance>(`Multiple command sets with ${args.options.title || args.options.clientSideComponentId} were found.`, resultAsKeyValuePair);
      return result.Id;
    }

    return listItemInstances[0].Id;
  }
}

export default new SpoTenantCommandSetRemoveCommand();