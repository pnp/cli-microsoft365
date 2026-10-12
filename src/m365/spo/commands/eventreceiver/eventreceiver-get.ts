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
  id: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `${e.input} is not a valid GUID.`
  }).optional().alias('i'),
  name: z.string().optional().alias('n'),
  scope: z.enum(['web', 'site']).optional().alias('s')
});

export type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoEventreceiverGetCommand extends SpoCommand {
  public get name(): string {
    return commands.EVENTRECEIVER_GET;
  }

  public get description(): string {
    return 'Gets a specific event receiver attached to the web, site or list (if any of the list options are filled in) by receiver name of id';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      })
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
    try {
      const eventReceiver: EventReceiver = await this.getEventReceiver(args);

      await logger.log(eventReceiver);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getEventReceiver(args: CommandArgs): Promise<EventReceiver> {
    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    if (!args.options.scope || args.options.scope === 'web') {
      requestOptions.url += `/web`;
    }
    else {
      requestOptions.url += '/site';
    }

    if (args.options.listId) {
      requestOptions.url += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      requestOptions.url += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestOptions.url += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    requestOptions.url += '/eventreceivers';


    if (args.options.id) {
      requestOptions.url += `(guid'${args.options.id}')`;

      const res = await request.get<EventReceiver>(requestOptions);

      // Instead of failing, the request will return an 'odata.null' object if the event receiver doesn't exist.
      if (res && (res as any)['odata.null']) {
        throw `The specified event receiver '${args.options.id}' does not exist.`;
      }

      return res;
    }
    else {
      requestOptions.url += `?$filter=receivername eq '${args.options.name}'`;

      const res = await request.get<{ value: EventReceiver[] }>(requestOptions);

      if (res.value.length === 0) {
        throw `The specified event receiver '${args.options.name}' does not exist.`;
      }

      if (res.value && res.value.length > 1) {
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('ReceiverId', res.value);
        return await cli.handleMultipleResultsFound<EventReceiver>(`Multiple event receivers with name '${args.options.name}' found.`, resultAsKeyValuePair);
      }

      return res.value[0];
    }
  }
}

export default new SpoEventreceiverGetCommand(); 