import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

const listTypes = ['List', 'Library', 'SitePages'] as const;
const locations = ['ContextMenu', 'CommandBar', 'Both'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  newTitle: z.string().optional().alias('t'),
  listType: z.enum(listTypes).optional().alias('l'),
  clientSideComponentId: z.string().refine(val => validation.isValidGuid(val), { message: 'clientSideComponentId is not a valid GUID' }).optional().alias('c'),
  clientSideComponentProperties: z.string().optional().alias('p'),
  webTemplate: z.string().optional().alias('w'),
  location: z.enum(locations).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTenantCommandSetSetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_COMMANDSET_SET;
  }

  public get description(): string {
    return 'Updates a ListView Command Set that is installed tenant wide.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => opts.newTitle || opts.listType || opts.clientSideComponentId || opts.clientSideComponentProperties || opts.webTemplate || opts.location, {
      error: 'Specify at least one property to update'
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);

      if (!appCatalogUrl) {
        throw 'No app catalog URL found';
      }

      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
      const listItem = await this.getListItemById(logger, appCatalogUrl, listServerRelativeUrl, args.options.id);

      if (listItem.TenantWideExtensionLocation.indexOf("ClientSideExtension.ListViewCommandSet") === -1) {
        throw 'The item is not a ListViewCommandSet';
      }

      await this.updateTenantWideExtension(appCatalogUrl, args.options, listServerRelativeUrl, logger);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListItemById(logger: Logger, webUrl: string, listServerRelativeUrl: string, id: string): Promise<any> {
    if (this.verbose) {
      await logger.logToStderr(`Getting the list item by id ${id}`);
    }
    const reqOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Items(${id})`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return await request.get<any>(reqOptions);
  }

  private async updateTenantWideExtension(appCatalogUrl: string, options: Options, listServerRelativeUrl: string, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Updating tenant wide extension to the TenantWideExtensions list');
    }

    const formValues: any = [];
    if (options.newTitle !== undefined) {
      formValues.push({
        FieldName: 'Title',
        FieldValue: options.newTitle
      });
    }

    if (options.clientSideComponentId !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionComponentId',
        FieldValue: options.clientSideComponentId
      });
    }

    if (options.location !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionLocation',
        FieldValue: this.getLocation(options.location)
      });
    }

    if (options.listType !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionListTemplate',
        FieldValue: this.getListTemplate(options.listType)
      });
    }

    if (options.clientSideComponentProperties !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionComponentProperties',
        FieldValue: options.clientSideComponentProperties
      });
    }

    if (options.webTemplate !== undefined) {
      formValues.push({
        FieldName: 'TenantWideExtensionWebTemplate',
        FieldValue: options.webTemplate
      });
    }

    const requestOptions: CliRequestOptions = {
      url: `${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Items(${options.id})/ValidateUpdateListItem()`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: {
        formValues: formValues
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private getLocation(location: string | undefined): string {
    switch (location) {
      case 'Both':
        return 'ClientSideExtension.ListViewCommandSet';
      case 'ContextMenu':
        return 'ClientSideExtension.ListViewCommandSet.ContextMenu';
      default:
        return 'ClientSideExtension.ListViewCommandSet.CommandBar';
    }
  }

  private getListTemplate(listTemplate: string | undefined): string {
    switch (listTemplate) {
      case 'SitePages':
        return '119';
      case 'Library':
        return '101';
      default:
        return '100';
    }
  }
}

export default new SpoTenantCommandSetSetCommand();