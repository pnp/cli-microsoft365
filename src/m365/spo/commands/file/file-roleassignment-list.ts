import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
}

class SpoFileRoleAssignmentListCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ROLEASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Lists all role assignments from a specific file.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
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
        option: '--fileId [fileId]'
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

  #initTypes(): void {
    this.types.string.push('webUrl', 'fileUrl', 'fileId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving role assignments for file with ${args.options.fileId || args.options.fileUrl} in site at ${args.options.webUrl}...`);
    }

    try {
      let file;
      if (args.options.fileId) {
        file = await spo.getFileById(args.options.webUrl, args.options.fileId, logger, this.verbose);
      }
      else {
        file = await spo.getFileByUrl(args.options.webUrl, args.options.fileUrl!, logger, this.verbose);
      }

      const fileRoleAssignments = await spo.getFileRoleAssignments(args.options.webUrl, file.ServerRelativeUrl, logger, this.verbose);
      await logger.log(fileRoleAssignments);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileRoleAssignmentListCommand();