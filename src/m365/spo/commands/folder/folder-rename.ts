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
  url: string;
  name: string;
}

class SpoFolderRenameCommand extends SpoCommand {

  public get name(): string {
    return commands.FOLDER_RENAME;
  }

  public get description(): string {
    return 'Renames a folder';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
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
        option: '-n, --name <name>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
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
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderRenameCommand();