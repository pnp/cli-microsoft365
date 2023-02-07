import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  id: string;
}

class SpoFileSharingLinkGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_GET;
  }

  public get description(): string {
    return 'Gets details about a specific sharing link of a file';
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
      },
      {
        option: '-i, --id <id>'
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
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving sharing link for file ${args.options.fileUrl || args.options.fileId} with id ${args.options.id}...`);
    }

    try {
      const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
      const requestOptions: CliRequestOptions = {
        url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFileSharingLinkGetCommand();
