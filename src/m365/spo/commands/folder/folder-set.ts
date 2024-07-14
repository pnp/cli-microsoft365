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
  color?: string;
}

class SpoFolderSetCommand extends SpoCommand {

  public get name(): string {
    return commands.FOLDER_SET;
  }

  public get description(): string {
    return 'Updates a folder';
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
        option: '--color [color]',
        autocomplete: Object.keys(FolderColorValues)
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
          return `Specify at least one of the options: name or color.`;
        }

        if (args.options.color && !Object.entries(FolderColorValues).flat().includes(args.options.color)) {
          return `'${args.options.color}' is not a valid value for option 'color'. Allowed values are ${Object.keys(FolderColorValues).join(', ')}, ${Object.values(FolderColorValues).join(', ')}.`;
        }

        return true;
      });
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'url', 'name', 'color');
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Updating folder '${args.options.name}'...`);
      }

      const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      if (!args.options.color) {
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
          url: `${args.options.webUrl}/_api/foldercoloring/${args.options.name ? 'renamefolder' : 'stampcolor'}(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            coloringInformation: {
              ColorHex: FolderColorValues[args.options.color!] || args.options.color
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