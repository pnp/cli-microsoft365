import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  vivaConnectionsDefaultStart?: boolean;
}

class SpoHomeSiteSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_SET;
  }

  public get description(): string {
    return 'Sets the specified site as the Home Site';
  }

  constructor() {
    super();
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        vivaConnectionsDefaultStart: typeof args.options.vivaConnectionsDefaultStart !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '--vivaConnectionsDefaultStart [vivaConnectionsDefaultStart]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  #initTypes(): void {
    this.types.boolean.push('vivaConnectionsDefaultStart');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
    try {
      let requestUrl: string = "";
      let requestBody: any;

      if (args.options.vivaConnectionsDefaultStart) {
        requestUrl = `${spoAdminUrl}/_api/SPO.Tenant/SetSPHSiteWithConfiguration`;
        requestBody = { sphSiteUrl: args.options.siteUrl, configuration: { vivaConnectionsDefaultStart: args.options.vivaConnectionsDefaultStart } };
      }
      else {
        requestUrl = `${spoAdminUrl}/_api/SPO.Tenant/SetSPHSite`;
        requestBody = { sphSiteUrl: args.options.siteUrl };
      }
      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-Type': 'application/json'
        },
        data: requestBody,
        responseType: 'json'
      };
      const res = await request.post<{ value: string; }>(requestOptions);
      logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoHomeSiteSetCommand();