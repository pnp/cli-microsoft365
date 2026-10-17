import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
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
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('l')
    .optional(),
  listTitle: z.string().alias('t').optional(),
  listUrl: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListRetentionLabelGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_RETENTIONLABEL_GET;
  }

  public get description(): string {
    return 'Gets the default retention label set on the specified list or library.';
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
      if (this.verbose) {
        await logger.logToStderr(`Getting label set on the list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
      }

      let listServerRelativeUrl: string = '';

      if (args.options.listUrl) {
        if (this.debug) {
          await logger.logToStderr(`Retrieving List from URL '${args.options.listUrl}'...`);
        }

        listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      }
      else {
        let requestUrl: string = `${args.options.webUrl}/_api/web/`;

        if (args.options.listId) {
          if (this.debug) {
            await logger.logToStderr(`Retrieving List from Id '${args.options.listId}'...`);
          }

          requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')?$expand=RootFolder&$select=RootFolder`;
        }
        else if (args.options.listTitle) {
          if (this.debug) {
            await logger.logToStderr(`Retrieving List from Title '${args.options.listTitle}'...`);
          }

          requestUrl += `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')?$expand=RootFolder&$select=RootFolder`;
        }

        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const listInstance: ListInstance = await request.get<ListInstance>(requestOptions);
        listServerRelativeUrl = listInstance.RootFolder.ServerRelativeUrl;
      }

      const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
      const reqOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          listUrl: listAbsoluteUrl
        }
      };

      const res = await request.post<any>(reqOptions);
      if (res['odata.null'] !== true) {
        await logger.log(res);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListRetentionLabelGetCommand();