import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  designPackageId?: string;
  designPackage?: string;
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
      },
      {
        option: '-p, --designPackage [designPackage]',
        autocomplete: ["Topic", "Showcase", "Blank"]
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.designPackageId &&
          !validation.isValidGuid(args.options.designPackageId)) {
          return `${args.options.designPackageId} is not a valid GUID.`;
        }

        if (args.options.designPackage) {
          if (['Topic', 'Showcase', 'Blank'].indexOf(args.options.designPackage) === -1) {
            return `${args.options.designPackage} is not a valid designPackage. Allowed values are Topic|Showcase|Blank`;
          }
        }

        if (args.options.designPackageId && args.options.designPackage) {
          return 'Specify designPackageId or designPackage but not both.';
        }

        return validation.isValidSharePointUrl(args.options.url);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const designPackageId = this.getDesignPackageId(args.options);

    if (this.verbose) {
      logger.logToStderr(`Enabling communication site with design package '${designPackageId}' at '${args.options.url}'...`);
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

  private getDesignPackageId(options: Options): string {
    if (options.designPackageId) {
      return options.designPackageId;
    }

    switch (options.designPackage) {
      case 'Blank':
        return 'f6cc5403-0d63-442e-96c0-285923709ffc';
      case 'Showcase':
        return '6142d2a0-63a5-4ba0-aede-d9fefca2c767';
      case 'Topic':
      default:
        return '96c933ac-3698-44c7-9f4a-5fd17d71af9e';
    }
  }
}

export default new SpoSiteCommSiteEnableCommand();