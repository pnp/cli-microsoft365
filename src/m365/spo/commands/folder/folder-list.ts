import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
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
  parentFolderUrl: string;
  recursive?: boolean;
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

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        recursive: !!args.options.recursive
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
        option: '--recursive [recursive]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving folders from site ${args.options.webUrl} parent folder ${args.options.parentFolderUrl} ${args.options.recursive ? '(recursive)' : ''}...`);
    }

    try {
      const resp = await this.getFolderList(args.options.webUrl, args.options.parentFolderUrl, args.options.recursive);
      logger.log(resp);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolderList(webUrl: string, parentFolderUrl: string, recursive?: boolean, folders: FolderProperties[] = []): Promise<FolderProperties[]> {
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, parentFolderUrl);


    const resp = await odata.getAllItems<FolderProperties>(`${webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')/folders`);
    if (resp.length > 0) {
      for (const folder of resp) {
        folders.push(folder);
        if (recursive) {
          await this.getFolderList(webUrl, folder.ServerRelativeUrl, recursive, folders);
        }
      }
    }

    return folders;
  }
}

module.exports = new SpoFolderListCommand();