import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FolderColorValues } from './FolderColor.js';
import { FolderProperties } from './FolderProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  parentFolderUrl: string;
  name: string;
  color?: string;
}

class SpoFolderAddCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ADD;
  }

  public get description(): string {
    return 'Creates a folder within a parent folder';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        color: typeof args.options.color !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-p, --parentFolderUrl <parentFolderUrl>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '--color [color]',
        autocomplete: Object.entries(FolderColorValues).flat()
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.color && !Object.entries(FolderColorValues).flat().includes(args.options.color)) {
          return `'${args.options.color}' is not a valid value for option 'color'. Allowed values are ${Object.keys(FolderColorValues).join(', ')}, ${Object.values(FolderColorValues).join(', ')}.`;
        }

        return true;
      });
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'parentFolderUrl', 'name', 'color');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding folder to site ${args.options.webUrl}...`);
    }

    const parentFolderServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.parentFolderUrl);
    const serverRelativeUrl: string = `${parentFolderServerRelativeUrl}/${args.options.name}`;

    const requestOptions: CliRequestOptions = {
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    if (args.options.color === undefined) {
      requestOptions.url = `${args.options.webUrl}/_api/web/folders/addUsingPath(decodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }
    else {
      requestOptions.url = `${args.options.webUrl}/_api/foldercoloring/createfolder(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}', overwrite=false)`;
      requestOptions.data = {
        coloringInformation: {
          ColorHex: FolderColorValues[args.options.color] || args.options.color
        }
      };
    }

    try {
      const folder = await request.post<FolderProperties>(requestOptions);
      await logger.log(folder);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderAddCommand();