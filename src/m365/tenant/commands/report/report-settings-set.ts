import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  hideUserInformation: boolean;
}

class TenantReportSettingsSetCommand extends GraphCommand {
  public get name(): string {
    return commands.REPORT_TENANTSETTINGS_SET;
  }

  public get description(): string {
    return 'Sets the tenant settings report';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      // Add unknown options to telemetry
      const unknownOptions = Object.keys(this.getUnknownOptions(args.options));
      const unknownOptionsObj = unknownOptions.reduce((obj, key) => ({ ...obj, [key]: true }), {});

      Object.assign(this.telemetryProperties, {
        hideUserInformation: typeof args.options.hideUserInformation !== 'undefined',
        ...unknownOptionsObj
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-h, --hideUserInformation  <hideUserInformation>',
        autocomplete: ['true', 'false']
      }
    );
  }


  #initTypes(): void {
    this.types.boolean.push('hideUserInformation');
  }


  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {

      const { hideUserInformation } = args.options;
      if (this.verbose) {
        await logger.logToStderr(`Updating report settings displayConcealedNames to ${hideUserInformation}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/reportSettings`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          'displayConcealedNames': args.options.hideUserInformation
        }
      };

      await request.patch(requestOptions);
    }
    catch (err) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantReportSettingsSetCommand();