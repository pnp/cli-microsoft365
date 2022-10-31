import { Logger } from '../../../../cli/Logger';
import { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FolderProperties } from './FolderProperties';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  url?: string;
  id?: string;
}

class SpoFolderGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_GET;
  }

  public get description(): string {
    return 'Gets information about the specified folder';
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
        id: typeof args.options.id !== 'undefined',
        url: typeof args.options.url !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-f, --url [url]'
      },
      {
        option: '-i, --id [id]'
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

        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['url', 'id']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving folder from site ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFolderById('${encodeURIComponent(args.options.id)}')`;
    }
    else if (args.options.url) {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      requestUrl = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const folder = await request.get<FolderProperties>(requestOptions);
      logger.log(folder);
    }
    catch (err: any) {
      if (err.statusCode && err.statusCode === 500) {
        throw new CommandError('Please check the folder URL. Folder might not exist on the specified URL');
      }

      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFolderGetCommand();