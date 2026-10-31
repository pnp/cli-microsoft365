import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

const expirationDateTimeMaxDays = 180;
const maxExpirationDateTime: Date = new Date();
// 180 days from now is the maximum expiration date for a webhook
maxExpirationDateTime.setDate(maxExpirationDateTime.getDate() + expirationDateTimeMaxDays);

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
  notificationUrl: z.string().alias('n'),
  expirationDateTime: z.string()
    .refine(date => !isNaN(Date.parse(date)), {
      error: `Provide the date in one of the following formats:
          'YYYY-MM-DD'
          'YYYY-MM-DDThh:mm'
          'YYYY-MM-DDThh:mmZ'
          'YYYY-MM-DDThh:mm±hh:mm'`
    })
    .refine(date => Date.parse(date) >= Date.now() && new Date(Date.parse(date)) < maxExpirationDateTime, {
      error: 'Provide an expiration date which is a date time in the future and within 6 months from now'
    })
    .alias('e')
    .optional(),
  clientState: z.string().alias('c').optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListWebhookAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_ADD;
  }

  public get description(): string {
    return 'Adds a new webhook to the specified list';
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
      await logger.logToStderr(`Adding webhook to list ${args.options.listId || args.options.listTitle || args.options.listUrl} located at site ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/Subscriptions`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions`;
    }

    const requestBody: any = {};
    requestBody.resource = args.options.listId || args.options.listTitle || args.options.listUrl;
    requestBody.notificationUrl = args.options.notificationUrl;
    // If no expiration date has been provided we will default to the
    // maximum expiration date of 180 days from now 
    requestBody.expirationDateTime = args.options.expirationDateTime
      ? new Date(args.options.expirationDateTime).toISOString()
      : maxExpirationDateTime.toISOString();
    if (args.options.clientState) {
      requestBody.clientState = args.options.clientState;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      const res = await request.post<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListWebhookAddCommand();