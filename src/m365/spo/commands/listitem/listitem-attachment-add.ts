import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import fs from 'fs';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} in option listId is not a valid GUID.`
  }).optional(),
  listTitle: z.string().optional(),
  listUrl: z.string().optional(),
  listItemId: z.number().int().positive(),
  filePath: z.string().refine(path => fs.existsSync(path), {
    error: e => `File with path '${e.input}' was not found.`
  }).alias('p'),
  fileName: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemAttachmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ATTACHMENT_ADD;
  }

  public get description(): string {
    return 'Adds an attachment to a list item';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: listId, listTitle, listUrl.',
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding an attachment to list item with id ${args.options.listItemId} on list ${args.options.listId || args.options.listTitle || args.options.listUrl} on web ${args.options.webUrl}.`);
    }

    try {
      const fileName = this.getFileName(args.options.filePath, args.options.fileName);
      const fileBody: Buffer = fs.readFileSync(args.options.filePath);
      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/${this.getListUrl(args.options.webUrl, args.options.listId, args.options.listTitle, args.options.listUrl)}/items(${args.options.listItemId})/AttachmentFiles/add(FileName='${fileName}')`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: fileBody,
        responseType: 'json'
      };

      const response = await request.post<any>(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      if (err.error &&
        err.error['odata.error'] &&
        err.error['odata.error'].message && err.error['odata.error'].message.value.indexOf('The document or folder name was not changed.') > -1) {
        this.handleError(err.error['odata.error'].message.value.split('\n')[0]);
      }
      else {
        this.handleRejectedODataJsonPromise(err);
      }
    }
  }

  private getFileName(filePath: string, fileName?: string): string {
    if (!fileName) {
      return filePath.replace(/^.*[\\/]/, '');
    }

    const extension = filePath.split('.').pop();
    if (!fileName.endsWith(`.${extension}`)) {
      fileName += `.${extension}`;
    }
    return fileName;
  }

  private getListUrl(webUrl: string, listId?: string, listTitle?: string, listUrl?: string): string {
    if (listId) {
      return `lists(guid'${formatting.encodeQueryParameter(listId)}')`;
    }
    else if (listTitle) {
      return `lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')`;
    }
    else {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl!);
      return `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }
  }
}

export default new SpoListItemAttachmentAddCommand();
