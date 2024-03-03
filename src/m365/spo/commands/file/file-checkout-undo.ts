import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  force?: boolean;
}

class SpoFileCheckoutUndoCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_CHECKOUT_UNDO;
  }

  public get description(): string {
    return 'Discards a checked out file';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        fileId: typeof args.options.fileId !== 'undefined',
        fileUrl: typeof args.options.fileUrl !== 'undefined',
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
        option: '--fileUrl [fileUrl]'
      },
      {
        option: '-i, --fileId [fileId]'
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

        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
          return `${args.options.fileId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'fileUrl', 'fileId');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const undoCheckout = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Undoing checkout for file ${args.options.fileId || args.options.fileUrl} on web ${args.options.webUrl}`);
        }

        let requestUrl: string = `${args.options.webUrl}/_api/web/`;

        if (args.options.fileId) {
          requestUrl += `getFileById('${args.options.fileId}')`;
        }
        else if (args.options.fileUrl) {
          const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
          requestUrl += `getFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
        }

        requestUrl += '/undocheckout';

        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await undoCheckout();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to undo the checkout for file ${args.options.fileId || args.options.fileUrl}?` });

      if (result) {
        await undoCheckout();
      }
    }
  }
}

export default new SpoFileCheckoutUndoCommand();