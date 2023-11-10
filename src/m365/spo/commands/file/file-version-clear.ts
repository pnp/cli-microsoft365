import { Cli } from '../../../../cli/Cli.js';
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

class SpoFileVersionClearCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_VERSION_CLEAR;
  }

  public get description(): string {
    return 'Deletes all file version history of a file in a SharePoint Document library';
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
        fileUrl: args.options.fileUrl,
        fileId: args.options.fileId,
        force: (!!args.options.force).toString()
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
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId as string)) {
          return `${args.options.fileId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Deletes all version history of the file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
    }

    try {
      if (args.options.force) {
        await this.clearVersions(args);
      }
      else {
        const result = await Cli.promptForConfirmation({ message: `Are you sure you want to delete all version history for file ${args.options.fileId || args.options.fileUrl}'?` });

        if (result) {
          await this.clearVersions(args);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async clearVersions(args: CommandArgs): Promise<void> {
    let requestUrl: string = `${args.options.webUrl}/_api/web/`;
    if (args.options.fileUrl) {
      const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
      requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/versions/DeleteAll()`;
    }
    else {
      requestUrl += `GetFileById('${args.options.fileId}')/versions/DeleteAll()`;
    }
    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }
}

export default new SpoFileVersionClearCommand();