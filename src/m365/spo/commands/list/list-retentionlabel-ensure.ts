import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListInstance } from './ListInstance.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `${e.input} is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  name: z.string(),
  listTitle: z.string().alias('t').optional(),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('l')
    .optional(),
  listUrl: z.string().optional(),
  syncToItems: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListRetentionLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_RETENTIONLABEL_ENSURE;
  }

  public get description(): string {
    return 'Sets a default retention label on the specified list or library.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: listId, listTitle, listUrl.',
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const listServerRelativeUrl: string = await this.getListServerRelativeUrl(args, logger);
      const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);

      await spo.applyDefaultRetentionLabelToList(args.options.webUrl, args.options.name, listAbsoluteUrl, args.options.syncToItems, logger, args.options.verbose);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListServerRelativeUrl(args: CommandArgs, logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr('Getting the list server relative URL');
    }

    if (args.options.listUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
    }

    let requestUrl = `${args.options.webUrl}/_api/web/`;

    if (args.options.listId) {
      requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
    }
    else {
      requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const listInstance = await request.get<ListInstance>(requestOptions);
    return listInstance.RootFolder.ServerRelativeUrl;
  }
}

export default new SpoListRetentionLabelEnsureCommand();