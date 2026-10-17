import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `${e.input} is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('l')
    .optional(),
  listTitle: z.string().alias('t').optional(),
  listUrl: z.string().optional(),
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('i'),
  force: z.boolean().alias('f').optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListWebhookRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified webhook from the list';
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
    const removeWebhook = async (): Promise<void> => {
      if (this.verbose) {
        const list: string = (args.options.listId || args.options.listId || args.options.listUrl) as string;
        await logger.logToStderr(`Webhook ${args.options.id} is about to be removed from list ${list} located at site ${args.options.webUrl}...`);
      }

      let requestUrl: string = `${args.options.webUrl}/_api/web`;

      if (args.options.listId) {
        requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
      }
      else if (args.options.listTitle) {
        requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        method: 'DELETE',
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
        // REST delete call doesn't return anything
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeWebhook();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove webhook ${args.options.id} from list ${args.options.listTitle || args.options.listId || args.options.listUrl} located at site ${args.options.webUrl}?` });

      if (result) {
        await removeWebhook();
      }
    }
  }
}

export default new SpoListWebhookRemoveCommand();