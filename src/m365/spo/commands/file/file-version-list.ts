import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
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
}

class SpoFileVersionListCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_VERSION_LIST;
  }

  public get description(): string {
    return 'Retrieves all versions of a file';
  }

  public defaultProperties(): string[] | undefined {
    return ['Created', 'ID', 'IsCurrentVersion', 'VersionLabel'];
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
        fileUrl: typeof args.options.fileUrl !== 'undefined',
        fileId: typeof args.options.fileId !== 'undefined'
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
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
      await logger.logToStderr(`Retrieving all versions of file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
    }

    try {
      let requestUrl = `${args.options.webUrl}/_api/web`;
      if (args.options.fileUrl) {
        const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
        requestUrl += `/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
      }
      else {
        requestUrl += `/GetFileById('${args.options.fileId}')`;
      }
      requestUrl += `/versions`;

      const response = await odata.getAllItems<any>(requestUrl);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileVersionListCommand();