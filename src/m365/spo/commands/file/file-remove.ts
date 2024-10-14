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

export interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  url?: string;
  recycle?: boolean;
  bypassSharedLock?: boolean;
  force?: boolean;
}

class SpoFileRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified file';
  }

  public alias(): string[] | undefined {
    return [commands.PAGE_TEMPLATE_REMOVE];
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
        id: typeof args.options.id !== 'undefined',
        url: typeof args.options.url !== 'undefined',
        recycle: !!args.options.recycle,
        bypassSharedLock: !!args.options.bypassSharedLock,
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
        option: '-i, --id [id]'
      },
      {
        option: '--url [url]'
      },
      {
        option: '--recycle'
      },
      {
        option: '--bypassSharedLock'
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

        if (args.options.id &&
          !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'url'] });
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'id', 'url');
    this.types.boolean.push('recycle', 'bypassSharedLock', 'force');
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeFile = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing file in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = '';

      if (args.options.id) {
        requestUrl = `${args.options.webUrl}/_api/web/GetFileById(guid'${formatting.encodeQueryParameter(args.options.id as string)}')`;
      }
      else {
        const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url!);
        requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
      }

      if (args.options.recycle) {
        requestUrl += `/recycle()`;
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.bypassSharedLock) {
        requestOptions.headers!.Prefer = 'bypass-shared-lock';
      }

      try {
        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeFile();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to ${args.options.recycle ? 'recycle' : 'remove'} the file ${args.options.id || args.options.url} located in site ${args.options.webUrl}?` });

      if (result) {
        await removeFile();
      }
    }
  }
}

export default new SpoFileRemoveCommand();