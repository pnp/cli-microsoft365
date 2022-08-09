import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  hideDefaultThemes: string;
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
    this.#initValidators();
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
        option: '--hideDefaultThemes <hideDefaultThemes>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.hideDefaultThemes !== 'false' &&
          args.options.hideDefaultThemes !== 'true') {
          return `${args.options.hideDefaultThemes} is not a valid boolean`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((spoAdminUrl: string): Promise<void> => {
        if (this.verbose) {
          logger.logToStderr(`Setting the value of the HideDefaultThemes setting to ${args.options.hideDefaultThemes}...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/thememanager/SetHideDefaultThemes`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          data: {
            hideDefaultThemes: args.options.hideDefaultThemes
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new SpoHideDefaultThemesSetCommand();