import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  hideDefaultThemes: boolean;
}

class SpoHideDefaultThemesSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HIDEDEFAULTTHEMES_SET;
  }

  public get description(): string {
    return 'Sets the value of the HideDefaultThemes setting';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        hideDefaultThemes: args.options.hideDefaultThemes
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--hideDefaultThemes <hideDefaultThemes>',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('hideDefaultThemes');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      if (this.verbose) {
        await logger.logToStderr(`Setting the value of the HideDefaultThemes setting to ${args.options.hideDefaultThemes}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/thememanager/SetHideDefaultThemes`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          hideDefaultThemes: args.options.hideDefaultThemes
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoHideDefaultThemesSetCommand();