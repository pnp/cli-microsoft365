import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FileProperties } from './FileProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  clearExistingPermissions?: boolean;
  force?: boolean;
}

class SpoFileRoleInheritanceBreakCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ROLEINHERITANCE_BREAK;
  }

  public get description(): string {
    return 'Breaks inheritance of a file. Keeping existing permissions is the default behavior.';
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
        fileId: typeof args.options.fileId !== 'undefined',
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
        option: '--fileUrl [fileUrl]'
      },
      {
        option: 'i, --fileId [fileId]'
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
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
          return `${args.options.fileId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const breakFileRoleInheritance = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Breaking role inheritance for file ${args.options.fileId || args.options.fileUrl}`);
      }
      try {
        const fileURL: string = await this.getFileURL(args);

        const keepExistingPermissions: boolean = !args.options.clearExistingPermissions;

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileURL)}')/ListItemAllFields/breakroleinheritance(${keepExistingPermissions})`,
          headers: {
            accept: 'application/json;odata=nometadata'
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
      await breakFileRoleInheritance();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to break the role inheritance of file ${args.options.fileUrl || args.options.fileId} located in site ${args.options.webUrl}?`
      });

      if (result.continue) {
        await breakFileRoleInheritance();
      }
    }
  }

  private async getFileURL(args: CommandArgs): Promise<string> {
    if (args.options.fileUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
    }

    const file: FileProperties = await spo.getFileById(args.options.webUrl, args.options.fileId!);
    return file.ServerRelativeUrl;
  }
}

export default new SpoFileRoleInheritanceBreakCommand();
