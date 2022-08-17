import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FolderProperties } from './FolderProperties';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  parentFolderUrl: string;
}

class SpoFolderListCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_LIST;
  }

  public get description(): string {
    return 'Returns all folders under the specified parent folder';
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'ServerRelativeUrl'];
  }

  constructor() {
    super();
  
    this.#initOptions();
    this.#initValidators();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-p, --parentFolderUrl <parentFolderUrl>'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving folders from site ${args.options.webUrl} parent folder ${args.options.parentFolderUrl}...`);
    }

    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.parentFolderUrl);
    const requestUrl: string = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/folders`;
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ value: FolderProperties[] }>(requestOptions)
      .then((resp: { value: FolderProperties[] }): void => {
        logger.log(resp.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoFolderListCommand();