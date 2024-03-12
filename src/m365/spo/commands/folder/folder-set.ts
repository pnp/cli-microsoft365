import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FolderColorValues } from './FolderColor.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  url: string;
  name?: string;
  color?: number | string;
}

class SpoFolderSetCommand extends SpoCommand {

  public get name(): string {
    return commands.FOLDER_SET;
  }

  public get description(): string {
    return 'Updates a folder';
  }

  public alias(): string[] | undefined {
    return [commands.FOLDER_RENAME];
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
        name: typeof args.options.name !== 'undefined',
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
        option: '--url <url>'
      },
      {
        option: '-n, --name [name]'
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

        if (args.options.color === undefined && args.options.name === undefined) {
          return `Specify atleast one of the options: name or color.`;
        }

        if (args.options.color !== undefined) {
          if (typeof args.options.color === 'number') {
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
    this.types.string.push('webUrl', 'url', 'name');
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Renaming folder ${args.options.url} to ${args.options.name}`);
      }

      const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      if (args.options.name && !args.options.color) {
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/Web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/ListItemAllFields`,
          headers: {
            accept: 'application/json;odata=nometadata',
            'if-match': '*'
          },
          data: {
            FileLeafRef: args.options.name,
            Title: args.options.name
          },
          responseType: 'json'
        };

        const response = await request.patch<any>(requestOptions);
        if (response && response['odata.null'] === true) {
          throw 'Folder not found.';
        }
      }
      else {
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/foldercoloring/${args.options.name !== undefined ? 'renamefolder' : 'stampcolor'}(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            coloringInformation: {
              ColorHex: `${typeof args.options.color === 'number' ? args.options.color : FolderColorValues[args.options.color!]}`
            },
            newName: args.options.name
          }
        };
        await request.post(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderSetCommand();