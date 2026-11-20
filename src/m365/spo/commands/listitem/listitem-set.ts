import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { basic } from '../../../../utils/basic.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.object({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.string().alias('i'),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} in option listId is not a valid GUID`
    })
    .optional()
    .alias('l'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  contentType: z.string().optional().alias('c'),
  systemUpdate: z.boolean().optional().alias('s')
}).catchall(z.unknown());

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemSetCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_SET;
  }

  public get description(): string {
    return 'Updates a list item in the specified list';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Specify either listId, listTitle or listUrl.',
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let contentTypeName: string = '';

    try {
      let requestUrl = `${args.options.webUrl}/_api/web`;

      if (args.options.listId) {
        requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else if (args.options.listTitle) {
        requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      if (args.options.contentType) {
        if (this.verbose) {
          await logger.logToStderr(`Getting content types for list...`);
        }

        const requestOptions: any = {
          url: `${requestUrl}/contenttypes?$select=Name,Id`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const contentTypes: any = await request.get(requestOptions);

        if (this.debug) {
          await logger.logToStderr('content type lookup response...');
          await logger.logToStderr(contentTypes);
        }

        const foundContentType: { Name: string; }[] = await basic.asyncFilter(contentTypes.value, async (ct: any) => {
          const contentTypeMatch: boolean = ct.Id.StringValue === args.options.contentType || ct.Name === args.options.contentType;

          if (this.debug) {
            await logger.logToStderr(`Checking content type value [${ct.Name}]: ${contentTypeMatch}`);
          }

          return contentTypeMatch;
        });

        if (this.debug) {
          await logger.logToStderr('content type filter output...');
          await logger.logToStderr(foundContentType);
        }

        if (foundContentType.length > 0) {
          contentTypeName = foundContentType[0].Name;
        }

        // After checking for content types, throw an error if the name is blank
        if (!contentTypeName || contentTypeName === '') {
          throw `Specified content type '${args.options.contentType}' doesn't exist on the target list`;
        }

        if (this.debug) {
          await logger.logToStderr(`using content type name: ${contentTypeName}`);
        }
      }

      const properties = this.mapRequestBody(args.options);

      const item = args.options.systemUpdate ?
        await spo.systemUpdateListItem(requestUrl, +args.options.id, logger, this.verbose, properties, contentTypeName)
        : await spo.updateListItem(requestUrl, args.options.id, properties, contentTypeName);
      delete item.ID;
      await logger.log(item);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const filteredData: { [key: string]: any } = {};
    const excludeOptions: string[] = [
      'listTitle',
      'listId',
      'listUrl',
      'webUrl',
      'id',
      'contentType',
      'systemUpdate',
      'debug',
      'verbose',
      'output',
      's',
      'i',
      'o',
      'u',
      't',
      '_'
    ];

    for (const key of Object.keys(options)) {
      if (!excludeOptions.includes(key)) {
        filteredData[key] = options[key];
      }
    }

    return filteredData;
  }
}

export default new SpoListItemSetCommand();