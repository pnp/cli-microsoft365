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
  color?: number | string;
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
        option: '--color [color]'
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

        if (args.options.color !== undefined) {
          if (typeof args.options.color === "number") {
            if (isNaN(args.options.color) || args.options.color < 0 || args.options.color > 15 || !Number.isInteger(args.options.color)) {
              return 'color should be an integer between 0 and 15.';
            }
          }
          else if (FolderColorValues[args.options.color] === undefined) {
            return `${args.options.color} is not a valid color value. Allowed values are ${Object.keys(FolderColorValues).join(', ')}.`;
          }
        }


        return true;
      });
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'parentFolderUrl', 'name');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding folder to site ${args.options.webUrl}...`);
    }

    const parentFolderServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.parentFolderUrl);
    const serverRelativeUrl: string = `${parentFolderServerRelativeUrl}/${args.options.name}`;

    const requestUrl: string = args.options.color !== undefined
      ? `${args.options.webUrl}/_api/foldercoloring/createfolder(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}', overwrite=false)`
      : `${args.options.webUrl}/_api/web/folders/addUsingPath(decodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    if (args.options.color !== undefined) {
      requestOptions.data = {
        'coloringInformation': {
          'ColorHex': `${typeof args.options.color === 'number' ? args.options.color : FolderColorValues[args.options.color]}`
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