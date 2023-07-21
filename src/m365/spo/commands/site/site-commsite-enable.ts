import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  designPackageId?: string;
  url: string;
}

class SpoSiteCommSiteEnableCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_COMMSITE_ENABLE;
  }

  public get description(): string {
    return 'Enables communication site features on the specified site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        designPackageId: typeof args.options.designPackageId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '-i, --designPackageId [designPackageId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.designPackageId &&
          !validation.isValidGuid(args.options.designPackageId)) {
          return `${args.options.designPackageId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.url);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const designPackageId: string = args.options.designPackageId || '96c933ac-3698-44c7-9f4a-5fd17d71af9e';

    if (this.verbose) {
      logger.logToStderr(`Enabling communication site at ${args.options.url}...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.url}/_api/sitepages/communicationsite/enable`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: { designPackageId },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteCommSiteEnableCommand();