import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  confirm?: boolean;
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
        confirm: (!!args.options.confirm).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-w, --webUrl <webUrl>'
      },
      {
        option: '-u, --fileUrl [fileUrl]'
      },
      {
        option: '-i, --fileId [fileId]'
      },
      {
        option: '--confirm'
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
    this.optionSets.push(['fileUrl', 'fileId']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Deletes all version history of the file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
    }

    try {
      if (args.options.confirm) {
        await this.clearVersions(args);
      }
      else {
        const result = await Cli.prompt<{ continue: boolean }>({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to delete all version history for file ${args.options.fileId || args.options.fileUrl}'?`
        });

        if (result.continue) {
          await this.clearVersions(args);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async clearVersions(args: CommandArgs): Promise<void> {
    let requestUrl;
    if (args.options.fileUrl) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(args.options.fileUrl)}')/versions/DeleteAll()`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${args.options.fileId}')/versions/DeleteAll()`;
    }
    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }
}

module.exports = new SpoFileVersionClearCommand();