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
  siteUrl: string;
  vivaConnectionsDefaultStart?: boolean;
}

class SpoTenantHomeSiteSetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_HOMESITE_SET;
  }

  public get description(): string {
    return 'Sets the specified site as the Home Site';
  }

  public alias(): string[] {
    return ['spo homesite set'];
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
    await this.showDeprecationWarning(logger, this.alias()[0], this.getCommandName());
    try {
      if (this.verbose) {
        await logger.logToStderr(`Setting the SharePoint home site to: ${args.options.siteUrl}...`);
        await logger.logToStderr('Attempting to retrieve the SharePoint admin URL.');
      }

      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPO.Tenant`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-Type': 'application/json'
        },
        responseType: 'json',
        data: {
          sphSiteUrl: args.options.siteUrl
        }
      };

      if (args.options.vivaConnectionsDefaultStart !== undefined) {
        requestOptions.url += '/SetSPHSiteWithConfiguration';
        requestOptions.data.configuration = { vivaConnectionsDefaultStart: args.options.vivaConnectionsDefaultStart };
      }
      else {
        requestOptions.url += '/SetSPHSite';
      }

      const res = await request.post<{ value: string; }>(requestOptions);
      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantHomeSiteSetCommand();