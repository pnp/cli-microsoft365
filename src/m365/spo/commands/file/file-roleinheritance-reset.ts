import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoFileGetCommand, { Options as SpoFileGetCommandOptions } from './file-get.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  force?: boolean;
}

class SpoFileRoleInheritanceResetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ROLEINHERITANCE_RESET;
  }

  public get description(): string {
    return 'Restores the role inheritance of a file';
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
    const resetFileRoleInheritance = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Resetting role inheritance for file ${args.options.fileId || args.options.fileUrl}`);
      }
      try {
        const fileURL: string = await this.getFileURL(args);

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(fileURL)}')/ListItemAllFields/resetroleinheritance`,
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
      await resetFileRoleInheritance();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to reset the role inheritance of file ${args.options.fileUrl || args.options.fileId} located in site ${args.options.webUrl}?`
      });

      if (result.continue) {
        await resetFileRoleInheritance();
      }
    }
  }

  private async getFileURL(args: CommandArgs): Promise<string> {
    if (args.options.fileUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
    }

    const options: SpoFileGetCommandOptions = {
      webUrl: args.options.webUrl,
      id: args.options.fileId,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(spoFileGetCommand as Command, { options: { ...options, _: [] } });
    const getFileOutput = JSON.parse(output.stdout);
    return getFileOutput.ServerRelativeUrl;
  }
}

export default new SpoFileRoleInheritanceResetCommand();
