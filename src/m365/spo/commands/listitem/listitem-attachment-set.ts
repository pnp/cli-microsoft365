import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import fs from 'fs';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  listItemId: number;
  fileName: string;
  filePath: string;
}

class SpoListItemAttachmentSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ATTACHMENT_SET;
  }

  public get description(): string {
    return 'Updates an attachment from a list item';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--listItemId <listItemId>'
      },
      {
        option: '-p, --filePath <filePath>'
      },
      {
        option: '-n, --fileName <fileName>'
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

        if (isNaN(args.options.listItemId)) {
          return `${args.options.listItemId} in option listItemId is not a valid number.`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID.`;
        }

        if (!fs.existsSync(args.options.filePath)) {
          return `File with path '${args.options.filePath}' was not found.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating attachment ${args.options.fileName} at path ${args.options.filePath} for list item with id ${args.options.listItemId} on list ${args.options.listId || args.options.listTitle || args.options.listUrl} on web ${args.options.webUrl}.`);
    }

    try {
      const fileName = this.getFileName(args.options.filePath, args.options.fileName);
      const fileBody: Buffer = fs.readFileSync(args.options.filePath);
      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/${this.getListUrl(args.options.webUrl, args.options.listId, args.options.listTitle, args.options.listUrl)}/items(${args.options.listItemId})/AttachmentFiles('${fileName}')/$value`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: fileBody,
        responseType: 'json'
      };

      await request.put(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getFileName(filePath: string, fileName: string): string {
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

export default new SpoListItemAttachmentSetCommand();
