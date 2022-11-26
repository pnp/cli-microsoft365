import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import { AxiosRequestConfig } from 'axios';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { formatting } from '../../../../utils/formatting';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as SpoFileGetCommand from './file-get';
import { Options as SpoFileGetCommandOptions } from './file-get';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  confirm?: boolean;
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
        confirm: !!args.options.confirm
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
        option: '--confirm'
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
    const resetFileRoleInheritance: () => Promise<void> = async (): Promise<void> => {
      try {
        const fileURL: string = await this.getFileURL(args);

        const requestOptions: AxiosRequestConfig = {
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

    if (args.options.confirm) {
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
      return args.options.fileUrl;
    }

    const options: SpoFileGetCommandOptions = {
      webUrl: args.options.webUrl,
      id: args.options.fileId,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(SpoFileGetCommand as Command, { options: { ...options, _: [] } });
    const getFileOutput = JSON.parse(output.stdout);
    return getFileOutput.ServerRelativeUrl;
  }
}

module.exports = new SpoFileRoleInheritanceResetCommand();
