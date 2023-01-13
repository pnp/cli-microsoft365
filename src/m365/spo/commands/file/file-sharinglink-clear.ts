import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FileSharingLinkUtil } from './FileSharingLinkUtil';
import { Options as SpoFileSharingLinkListOptions } from './file-sharinglink-list';
import { Cli } from '../../../../cli/Cli';
import request, { CliRequestOptions } from '../../../../request';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  id: string;
}

class SpoFileSharingLinkClearCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_CLEAR;
  }

  public get description(): string {
    return 'Removes all sharing links of a file';
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
        scope: typeof args.options.scope !== 'undefined',
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
        option: '--fileId [fileId]'
      },
      {
        option: '--scope [scope]',
        autocomplete: FileSharingLinkUtil.allowedScopes
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

        if (args.options.scope && FileSharingLinkUtil.allowedScopes.some(args.options.scope)) {
          return `'${args.options.scope}' is not a valid scope. Allowed values are: ${FileSharingLinkUtil.allowedScopes.join(',')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearSharingLinks: () => Promise<void> = async (): Promise<void> => {
      try {
        if (this.verbose) {
          logger.logToStderr(`Clearing sharing links for file ${args.options.fileUrl || args.options.fileId} ${args.options.scope && ` and scope ${args.options.scope}`}`);
        }

        const sharingLinkListOptions: SpoFileSharingLinkListOptions = {
          webUrl: args.options.webUrl,
          fileId: args.options.fileId,
          fileUrl: args.options.fileUrl,
          scope: args.options.scope,
          debug: this.debug,
          verbose: this.verbose
        };

        const commandOutput = await Cli.executeCommandWithOutput(getCommand as Command, { options: { ...sharingLinkListOptions, _: [] } });

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await clearSharingLinks();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove sharing link ${args.options.id} of file ${args.options.fileUrl || args.options.fileId}?`
      });

      if (result.continue) {
        await clearSharingLinks();
      }
    }
  }
}

module.exports = new SpoFileSharingLinkClearCommand();
