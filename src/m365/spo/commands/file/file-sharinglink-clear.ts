import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
  scope?: string;
  force?: boolean;
}

class SpoFileSharingLinkClearCommand extends SpoCommand {
  private readonly allowedScopes: string[] = ['anonymous', 'users', 'organization'];

  public get name(): string {
    return commands.FILE_SHARINGLINK_CLEAR;
  }

  public get description(): string {
    return 'Removes sharing links of a file';
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
        option: '--fileId [fileId]'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: this.allowedScopes
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

        if (args.options.scope && this.allowedScopes.indexOf(args.options.scope) === -1) {
          return `'${args.options.scope}' is not a valid scope. Allowed values are: ${this.allowedScopes.join(',')}`;
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
          await logger.logToStderr(`Clearing sharing links for file ${args.options.fileUrl || args.options.fileId}${args.options.scope ? ` with scope ${args.options.scope}` : ''}`);
        }

        const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
        const sharingLinks = await spo.getFileSharingLinks(args.options.webUrl, args.options.fileId, args.options.fileUrl, args.options.scope, logger, this.verbose);

        const requestOptions: CliRequestOptions = {
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        for (const sharingLink of sharingLinks) {
          requestOptions.url = `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${sharingLink.id}`;
          await request.delete(requestOptions);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await clearSharingLinks();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to clear the sharing links of file ${args.options.fileUrl || args.options.fileId}${args.options.scope ? ` with scope ${args.options.scope}` : ''}?`
      });

      if (result.continue) {
        await clearSharingLinks();
      }
    }
  }
}

export default new SpoFileSharingLinkClearCommand();
