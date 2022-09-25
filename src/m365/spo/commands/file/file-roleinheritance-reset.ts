import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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
    return 'Restores the role inheritance of list item, file, or folder';
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
        webUrl: typeof args.options.webUrl !== 'undefined',
        fileUrl: typeof args.options.fileUrl !== 'undefined',
        fileId: typeof args.options.fileId !== 'undefined'
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
    this.optionSets.push(['fileId', 'fileUrl']);
  }

  private getFileURL(args: CommandArgs): Promise<string> {
    if (args.options.fileUrl) {
      return Promise.resolve(args.options.fileUrl);
    }

    const options: SpoFileGetCommandOptions = {
      webUrl: args.options.webUrl,
      id: args.options.fileId,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    return Cli.executeCommandWithOutput(SpoFileGetCommand as Command, { options: { ...options, _: [] } })
      .then((output: CommandOutput): Promise<string> => {
        const getFileOutput = JSON.parse(output.stdout);
        return Promise.resolve(getFileOutput.ServerRelativeUrl);
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const resetFileRoleInheritance: () => Promise<void> = async (): Promise<void> => {
      try {
        const fileURL: string = await this.getFileURL(args);

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${fileURL}')/ListItemAllFields/resetroleinheritance`,
          headers: {
            accept: 'application/json;odata.metadata=none'
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
      resetFileRoleInheritance();
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
}

module.exports = new SpoFileRoleInheritanceResetCommand();
