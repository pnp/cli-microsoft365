import fs from 'fs';
import os from 'os';
import { v4 } from 'uuid';
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
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} in option listId is not a valid GUID.`
  }).optional().alias('l'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  filePath: z.string()
    .superRefine((path, ctx) => {
      if (!fs.existsSync(path)) {
        ctx.addIssue({ code: 'custom', message: `File with path ${path} does not exist.` });
        return;
      }

      const fileContent = fs.readFileSync(path, 'utf-8');
      const jsonContent: any[] = formatting.parseCsvToJson(fileContent);

      const idKey = Object.keys(jsonContent[0]).find(key => key.toLowerCase() === 'id');

      if (!idKey) {
        ctx.addIssue({ code: 'custom', message: `The file does not contain the required header row with the column name 'ID'.` });
        return;
      }

      const invalidIDs = validation.isValidPositiveIntegerArray(jsonContent.map(element => element[idKey].toString().trim()).join(','));

      if (invalidIDs !== true) {
        ctx.addIssue({ code: 'custom', message: `The file contains one or more invalid IDs: '${invalidIDs}'.` });
      }
    })
    .optional()
    .alias('p'),
  ids: z.string()
    .superRefine((val, ctx) => {
      const invalidIDs = validation.isValidPositiveIntegerArray(val);

      if (invalidIDs !== true) {
        ctx.addIssue({ code: 'custom', message: `Option 'ids' contains one or more invalid IDs: '${invalidIDs}'.` });
      }
    })
    .optional()
    .alias('i'),
  permanent: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemBatchRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_BATCH_REMOVE;
  }

  public get description(): string {
    return 'Removes items from a list in batch';
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
      })
      .refine(opts => [opts.filePath, opts.ids].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: filePath, ids.',
        params: {
          customCode: 'optionSet',
          options: ['filePath', 'ids']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeListItems = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr('Removing the listitems from SharePoint...');
        }

        let idsToRemove: string[] = [];

        if (args.options.filePath) {
          const csvContent = fs.readFileSync(args.options.filePath, 'utf-8');
          const jsonContent = formatting.parseCsvToJson(csvContent);

          const idKey = Object.keys(jsonContent[0]).find(key => key.toLowerCase() === 'id');

          idsToRemove = jsonContent.map((item: any) => item[idKey!]);
        }
        else {
          idsToRemove = formatting.splitAndTrim(args.options.ids!);
        }

        await this.removeItemsAsBatch(idsToRemove, args.options, logger);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeListItems();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to ${args.options.permanent ? 'permanently remove' : 'recycle'} the list items from list ${args.options.listId || args.options.listTitle || args.options.listUrl} located in site ${args.options.webUrl}?` });

      if (result) {
        await removeListItems();
      }
    }
  }

  private async removeItemsAsBatch(items: string[], options: Options, logger: Logger): Promise<void> {
    const itemsChunks: string[][] = this.getChunkedArray(items, 10);

    for (const [index, chunk] of itemsChunks.entries()) {
      if (this.verbose) {
        await logger.logToStderr(`Processing chunk ${index + 1} of ${itemsChunks.length}...`);
      }

      await this.postBatchData(chunk, options.webUrl, options);
    }
  }

  private async postBatchData(chunk: string[], webUrl: string, options: Options): Promise<void> {
    const batchId = v4();

    const requestBody = this.getRequestBody(chunk, batchId, options);
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/$batch`,
      headers: {
        'Content-Type': `multipart/mixed; boundary=batch_${batchId}`,
        'Accept': 'application/json;odata=verbose'
      },
      data: requestBody.join('')
    };
    const response: string = await request.post(requestOptions);
    const errors = this.parseBatchResponseBody(response, chunk);

    if (errors.length > 0) {
      throw `Creating some items failed with the following errors: ${os.EOL}${errors.map(error => { return `- ${error}`; }).join(os.EOL)}`;
    }
  }

  private getRequestBody(chunk: string[], batchId: string, options: Options): string[] {
    const changeSetId = v4();
    const batchBody: string[] = [];

    batchBody.push(`--batch_${batchId}\n`);
    batchBody.push(`Content-Type: multipart/mixed; boundary="changeset_${changeSetId}"\n\n`);
    batchBody.push('Content-Transfer-Encoding: binary\n\n');

    for (const item of chunk) {
      const actionUrl = this.getActionUrl(options, item);
      batchBody.push(`--changeset_${changeSetId}\n`);
      batchBody.push('Content-Type: application/http\n');
      batchBody.push('Content-Transfer-Encoding: binary\n\n');
      batchBody.push(`DELETE ${actionUrl} HTTP/1.1\n`);
      batchBody.push(`Accept: application/json;odata=nometadata\n`);
      batchBody.push(`If-Match: *\n\n`);
    }

    batchBody.push(`\n\n`);
    batchBody.push(`--changeset_${changeSetId}--\n\n`);
    batchBody.push(`--batch_${batchId}--\n`);

    return batchBody;
  }

  private parseBatchResponseBody(response: string, chunk: string[]): string[] {
    const errors: string[] = [];

    response.split('\r\n')
      .filter((line: string) => line.startsWith('{'))
      .forEach((line: string, index: number) => {
        const parsedResponse: any = JSON.parse(line);

        if (parsedResponse.error) {
          const error = parsedResponse.error as { message: { value: string } };
          errors.push(`Item ID ${chunk[index]}: ${error.message.value}`);
        }
      });

    return errors;
  };

  private getChunkedArray(inputArray: string[], chunkSize: number): string[][] {
    const result: string[][] = [];

    for (let i = 0; i < inputArray.length; i += chunkSize) {
      result.push(inputArray.slice(i, i + chunkSize));
    }

    return result;
  }

  private getActionUrl(options: Options, item: string): string {
    let requestUrl = '';

    if (options.listId) {
      requestUrl += `lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
    }
    else if (options.listTitle) {
      requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }
    else if (options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    requestUrl += `/items(${item})`;

    if (!options.permanent) {
      requestUrl += '/recycle()';
    }

    return requestUrl;
  }
}

export default new SpoListItemBatchRemoveCommand();