import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListItemInstance } from './ListItemInstance.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(webUrl => validation.isValidSharePointUrl(webUrl) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listId: z.string()
    .refine(listId => validation.isValidGuid(listId), {
      error: e => `${e.input} in option listId is not a valid GUID`
    })
    .optional()
    .alias('l'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  id: z.number().int().positive().optional().alias('i'),
  uniqueId: z.string()
    .refine(uniqueId => validation.isValidGuid(uniqueId), {
      error: e => `${e.input} in option uniqueId is not a valid GUID`
    })
    .optional(),
  fields: z.string().optional(),
  withPermissions: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_GET;
  }

  public get description(): string {
    return 'Gets a list item from the specified list';
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
      })
      .refine(opts => [opts.id, opts.uniqueId].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: id, uniqueId.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'uniqueId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
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

    const fieldsSelect: string[] = args.options.fields ? args.options.fields.split(',') : [];
    const fieldsWithSlash: string[] = fieldsSelect.filter(item => item.includes('/'));
    const fieldsToExpand: string[] = fieldsWithSlash.map(e => e.split('/')[0]);
    const expandFieldsArray: string[] = fieldsToExpand.filter((item, pos) => fieldsToExpand.indexOf(item) === pos);
    const fieldExpand: string = expandFieldsArray.length > 0 ? `&$expand=${expandFieldsArray.join(",")}` : ``;

    if (args.options.id) {
      requestUrl += `/items(${args.options.id})`;
    }
    else {
      requestUrl += `/GetItemByUniqueId(guid'${args.options.uniqueId}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$select=${formatting.encodeQueryParameter(fieldsSelect.join(","))}${fieldExpand}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const itemProperties = await request.get<any>(requestOptions);
      if (args.options.withPermissions) {
        requestOptions.url = `${requestUrl}/RoleAssignments?$expand=Member,RoleDefinitionBindings`;
        const response = await request.get<{ value: any[] }>(requestOptions);
        response.value.forEach((r: any) => {
          r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
        });
        itemProperties.RoleAssignments = response.value;
      }
      delete itemProperties['ID'];
      await logger.log(<ListItemInstance>itemProperties);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListItemGetCommand();
