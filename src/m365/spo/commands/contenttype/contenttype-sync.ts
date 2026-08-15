import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.string().optional().alias('i'),
  name: z.string().optional().alias('n'),
  listTitle: z.string().optional(),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    }).optional(),
  listUrl: z.string().optional()
});

export type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoContentTypeSyncCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_SYNC;
  }

  public get description(): string {
    return 'Adds a published content type from the content type hub to a site or syncs its latest changes';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      })
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(x => x !== undefined).length <= 1, {
        error: `Specify either 'listId', 'listTitle' or 'listUrl'.`,
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { listId, listTitle, listUrl, webUrl } = args.options;
    const url: URL = new URL(webUrl);
    const baseUrl = `${url.origin}/_api/v2.0/sites/`;

    try {
      const [siteId, webId] = await Promise.all([
        spo.getSiteIdBySPApi(webUrl, logger, this.verbose),
        spo.getWebId(webUrl, logger, this.verbose)
      ]);
      const siteUrl = `${url.host},${siteId},${webId}`;
      const targetListId = listId || (listTitle || listUrl ? await spo.getListId(webUrl, listTitle, listUrl, logger, this.verbose) : undefined);
      const listPath = targetListId ? `/lists/${targetListId}` : '';
      const contentTypeId = await this.getContentTypeId(webUrl, args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Adding or syncing the content type...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${baseUrl}${siteUrl}${listPath}/contenttypes/addCopyFromContentTypeHub`,
        headers: {
          'accept': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false'
        },
        responseType: 'json',
        data: {
          contentTypeId: contentTypeId
        }
      };

      const response = await request.post(requestOptions);

      // The endpoint only returns a response if the content type has been added for the first time
      // When syncing, the response will be an empty string, which should not be logged.
      if (typeof response === 'object') {
        await logger.log(response);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContentTypeId(webUrl: string, options: Options, logger: Logger): Promise<string> {
    if (options.id) {
      return options.id;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving content type Id by name...`);
    }

    const contentTypes: { StringId: string }[] = await odata.getAllItems(`${webUrl}/_api/web/AvailableContentTypes?$filter=Name eq '${formatting.encodeQueryParameter(options.name!)}'&$select=StringId,Name`);

    if (contentTypes.length === 0) {
      throw `Content type with name ${options.name} not found.`;
    }

    return contentTypes[0].StringId;
  }
}

export default new SpoContentTypeSyncCommand();