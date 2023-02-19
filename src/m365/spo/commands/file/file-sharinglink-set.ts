import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { spo } from '../../../../utils/spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id: string;
  fileId?: string;
  fileUrl?: string;
  expirationDateTime: string;
}

class SpoFileSharingLinkSetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_SET;
  }

  public get description(): string {
    return 'Updates a sharing link of a file';
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
        fileUrl: typeof args.options.fileUrl !== 'undefined'
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
        option: '--id <id>'
      },
      {
        option: '--expirationDateTime <expirationDateTime>'
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

        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (!validation.isValidISODateTime(args.options.expirationDateTime)) {
          return `'${args.options.expirationDateTime}' is not a valid ISO date string`;
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
      logger.logToStderr(`Updating sharing link of file ${args.options.fileId || args.options.fileUrl}...`);
    }

    try {
      const fileDetails = await spo.getVroomFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);

      const requestOptions: CliRequestOptions = {
        url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          expirationDateTime: args.options.expirationDateTime
        }
      };

      const sharingLink = await request.patch<any>(requestOptions);

      logger.log(sharingLink);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFileSharingLinkSetCommand();