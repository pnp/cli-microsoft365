import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.hideDefaultThemes = args.options.hideDefaultThemes;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
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
            hideDefaultThemes: args.options.hideDefaultThemes,
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--hideDefaultThemes <hideDefaultThemes>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.hideDefaultThemes !== 'false' &&
      args.options.hideDefaultThemes !== 'true') {
      return `${args.options.hideDefaultThemes} is not a valid boolean`;
    }

    return true;
  }
}

module.exports = new SpoHideDefaultThemesSetCommand();