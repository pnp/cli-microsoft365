import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListInstance } from '../list/ListInstance.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(webUrl => validation.isValidSharePointUrl(webUrl) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listId: z.string()
    .refine(listId => validation.isValidGuid(listId), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional(),
  listTitle: z.string().optional(),
  listUrl: z.string().optional(),
  listItemId: z.number().int().positive()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemRecordLockCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_RECORD_LOCK;
  }

  public get description(): string {
    return 'Locks the list item record';
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
    if (this.verbose) {
      await logger.logToStderr(`Locking the list item record ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
    }
    try {
      let listRestUrl: string = '';
      let listServerRelativeUrl: string = '';

      if (args.options.listUrl) {
        const listServerRelativeUrlFromPath: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);

        listServerRelativeUrl = listServerRelativeUrlFromPath;
      }
      else {
        if (args.options.listId) {
          listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
        }
        else {
          listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
        }

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/${listRestUrl}?$expand=RootFolder&$select=RootFolder`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const listInstance: ListInstance = await request.get<ListInstance>(requestOptions);
        listServerRelativeUrl = listInstance.RootFolder.ServerRelativeUrl;
      }

      const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
      const requestUrl: string = `${args.options.webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.LockRecordItem()`;
      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          listUrl: listAbsoluteUrl,
          itemId: args.options.listItemId
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListItemRecordLockCommand();