import { Cli } from '../../../../cli/Cli.js';
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
  folderUrl: string;
  clearExistingPermissions?: boolean;
  force?: boolean;
}

class SpoFolderRoleInheritanceBreakCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ROLEINHERITANCE_BREAK;
  }

  public get description(): string {
    return 'Breaks the role inheritance of a folder. Keeping existing permissions is the default behavior.';
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
        clearExistingPermissions: !!args.options.clearExistingPermissions,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--folderUrl <folderUrl>'
      },
      {
        option: '-c, --clearExistingPermissions'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const keepExistingPermissions: boolean = !args.options.clearExistingPermissions;
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
    const roleFolderUrl: string = urlUtil.getWebRelativePath(args.options.webUrl, args.options.folderUrl);
    let requestUrl: string = `${args.options.webUrl}/_api/web/`;

    const breakFolderRoleInheritance = async (): Promise<void> => {
      try {
        if (roleFolderUrl.split('/').length === 2) {
          requestUrl += `GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
        }
        else {
          requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields`;
        }
        const requestOptions: CliRequestOptions = {
          url: `${requestUrl}/breakroleinheritance(${keepExistingPermissions})`,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };
        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await breakFolderRoleInheritance();
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to break the role inheritance of folder ${args.options.folderUrl} located in site ${args.options.webUrl}?`);

      if (result) {
        await breakFolderRoleInheritance();
      }
    }
  }
}

export default new SpoFolderRoleInheritanceBreakCommand();