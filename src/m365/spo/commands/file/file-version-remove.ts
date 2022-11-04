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
  label: string;
  fileUrl?: string;
  fileId?: string;
  confirm?: boolean;
}

class SpoFileVersionRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_VERSION_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific version of a file in a SharePoint Document library';
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
        option: '--label <label>'
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

  #initTypes(): void {
    this.types.string.push('label');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removes version ${args.options.label} of the file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
    }

    try {
      if (args.options.confirm) {
        await this.removeVersion(args);
      }
      else {
        const result = await Cli.prompt<{ continue: boolean }>({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove the version ${args.options.label} from file ${args.options.fileId || args.options.fileUrl}'?`
        });

        if (result.continue) {
          await this.removeVersion(args);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeVersion(args: CommandArgs): Promise<void> {
    let requestUrl: string = `${args.options.webUrl}/_api/web/`;
    if (args.options.fileUrl) {
      requestUrl += `GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(args.options.fileUrl)}')/versions/DeleteByLabel('${args.options.label}')`;
    }
    else {
      requestUrl += `GetFileById('${args.options.fileId}')/versions/DeleteByLabel('${args.options.label}')`;
    }
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.delete(requestOptions);
  }
}

module.exports = new SpoFileVersionRemoveCommand();