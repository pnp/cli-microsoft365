import fs from 'fs';
import os from 'os';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { v4 } from 'uuid';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { cli } from '../../../../cli/cli.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  filePath?: string;
  ids?: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  recycle?: boolean;
  force?: boolean;
}

class SpoListItemBatchRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_BATCH_REMOVE;
  }

  public get description(): string {
    return 'Removes items from a list in batch';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        filePath: typeof args.options.filePath !== 'undefined',
        ids: typeof args.options.ids !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        recycle: !!args.options.recycle,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-p, --filePath [filePath]'
      },
      {
        option: '-i, --ids [ids]'
      },
      {
        option: '-r, --recycle'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID.`;
        }

        if (args.options.filePath) {

          if (!fs.existsSync(args.options.filePath)) {
            return `File with path ${args.options.filePath} does not exist.`;
          }
          // read and validate content
          const fileContent = fs.readFileSync(args.options.filePath, 'utf-8');
          const jsonContent: any[] = formatting.parseCsvToJson(fileContent);

          if (!jsonContent[0].hasOwnProperty('ID')) {
            return `The file does not contain the required header row with the column name 'ID'.`;
          }

          const nonNumbers = jsonContent.filter(element => isNaN(Number(element['ID'].toString().trim()))).map(element => element['ID']);

          if (nonNumbers.length > 0) {
            return `The specified ids '${nonNumbers.join(', ')}' are invalid numbers.`;
          }
        }

        if (args.options.ids) {
          const nonNumbers = formatting.splitAndTrim(args.options.ids).filter(element => isNaN(Number(element)));

          if (nonNumbers.length > 0) {
            return `The specified ids '${nonNumbers.join(', ')}' are invalid numbers.`;
          }
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push(
      'webUrl',
      'listId',
      'listTitle',
      'listUrl',
      'ids',
      'filePath'
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['listId', 'listTitle', 'listUrl']
      },
      {
        options: ['filePath', 'ids']
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeListItems = async (): Promise<void> => {
      try {

        if (this.verbose) {
          logger.logToStderr('Removing the listitems from SharePoint...');
        }

        let idsToRemove: string[] = [];

        if (args.options.filePath) {
          const csvContent = fs.readFileSync(args.options.filePath, 'utf-8');
          const jsonContent = formatting.parseCsvToJson(csvContent);
          idsToRemove = jsonContent.map((item: { ID: string }) => item['ID']);
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
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the list items from list ${args.options.listId || args.options.listTitle || args.options.listUrl} located in site ${args.options.webUrl}?` });

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

    if (options.recycle) {
      requestUrl += '/recycle()';
    }

    return requestUrl;
  }
}

export default new SpoListItemBatchRemoveCommand();