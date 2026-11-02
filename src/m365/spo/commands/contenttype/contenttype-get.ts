import { z } from 'zod';
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
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listTitle: z.string().optional().alias('l'),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    }).optional(),
  listUrl: z.string().optional(),
  id: z.string().optional().alias('i'),
  name: z.string().optional().alias('n')
});

export type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoContentTypeGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_GET;
  }

  public get description(): string {
    return 'Retrieves information about the specified list or site content type';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
      error: `Specify either 'id' or 'name', but not both.`,
      params: {
        customCode: 'optionSet',
        options: ['id', 'name']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let requestUrl: string = `${args.options.webUrl}/_api/web`;
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

    requestUrl += "/contenttypes";

    if (args.options.id) {
      requestUrl += `('${formatting.encodeQueryParameter(args.options.id)}')?$expand=Parent`;
    }
    else if (args.options.name) {
      requestUrl += `?$filter=Name eq '${formatting.encodeQueryParameter(args.options.name)}'&$expand=Parent`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      let res = await request.get<any>(requestOptions);
      let errorMessage: string = '';

      if (args.options.name) {
        if (res.value.length === 0) {
          errorMessage = `Content type with name ${args.options.name} not found`;
        }
        else {
          res = res.value[0];
        }
      }

      if (args.options.id && res['odata.null'] === true) {
        errorMessage = `Content type with ID ${args.options.id} not found`;
      }

      if (errorMessage) {
        throw errorMessage;
      }

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoContentTypeGetCommand();