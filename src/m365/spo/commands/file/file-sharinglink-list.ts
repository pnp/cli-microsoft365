import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  fileId?: string;
  fileUrl?: string;
  scope?: string;
}

class SpoFileSharingLinkListCommand extends SpoCommand {
  private static readonly allowedScopes: string[] = ['anonymous', 'users', 'organization'];

  public get name(): string {
    return commands.FILE_SHARINGLINK_LIST;
  }

  public get description(): string {
    return 'Lists all the sharing links of a specific file';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'scope', 'roles', 'link'];
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
        fileId: typeof args.options.fileId !== 'undefined',
        fileUrl: typeof args.options.fileUrl !== 'undefined',
        scope: typeof args.options.scope !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--fileId [fileId]'
      },
      {
        option: '--fileUrl [fileUrl]'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: SpoFileSharingLinkListCommand.allowedScopes
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

        if (args.options.scope && SpoFileSharingLinkListCommand.allowedScopes.indexOf(args.options.scope) === -1) {
          return `'${args.options.scope}' is not a valid scope. Allowed values are: ${SpoFileSharingLinkListCommand.allowedScopes.join(',')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving sharing links for file ${args.options.fileId || args.options.fileUrl}...`);
    }

    try {
      const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
      let url = `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions?$filter=Link ne null`;
      if (args.options.scope) {
        url += ` and Link/Scope eq '${args.options.scope}'`;
      }

      const sharingLinks = await odata.getAllItems<any>(url);

      if (!args.options.output || !Cli.shouldTrimOutput(args.options.output)) {
        await logger.log(sharingLinks);
      }
      else {
        //converted to text friendly output
        await logger.log(sharingLinks.map(i => {
          return {
            id: i.id,
            roles: i.roles.join(','),
            link: i.link.webUrl,
            scope: i.link.scope
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileSharingLinkListCommand();