import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { EventReceiver } from './EventReceiver.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listId: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `${e.input} is not a valid GUID.`
  }).optional(),
  listTitle: z.string().optional(),
  listUrl: z.string().optional(),
  scope: z.enum(['web', 'site']).optional().alias('s')
});

export type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoEventreceiverListCommand extends SpoCommand {
  public get name(): string {
    return commands.EVENTRECEIVER_LIST;
  }

  public get description(): string {
    return 'Retrieves event receivers for the specified web, site or list.';
  }

  public defaultProperties(): string[] | undefined {
    return ['ReceiverId', 'ReceiverName'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => {
        const listOptions = [opts.listId, opts.listTitle, opts.listUrl];
        return !(listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1);
      }, {
        message: `Specify either list id or title or list url`,
        params: {
          customCode: 'required'
        }
      })
      .refine(opts => !(opts.scope === 'site' && (opts.listId || opts.listUrl || opts.listTitle)), {
        message: 'Scope cannot be set to site when retrieving list event receivers.',
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let requestUrl = `${args.options.webUrl}/_api/`;
    let listUrl: string = '';

    if (args.options.listId) {
      listUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
    }
    else if (args.options.listTitle) {
      listUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      listUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
    }

    if (!args.options.scope || args.options.scope === 'web') {
      requestUrl += `web/${listUrl}eventreceivers`;
    }
    else {
      requestUrl += 'site/eventreceivers';
    }

    try {
      const res = await odata.getAllItems<EventReceiver>(requestUrl);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoEventreceiverListCommand();