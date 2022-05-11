import { Cli, Logger } from '../../../cli';
import type * as open from 'open';
import { validation } from '../../../utils/validation';
import { settingsNames } from '../../../settingsNames';
import AppCommand from '../../base/AppCommand';
import GlobalOptions from '../../../GlobalOptions';
import commands from '../commands';
import { CommandOption } from '../../../Command';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  preview?: boolean;
}

class AppOpenCommand extends AppCommand {
  private _open: typeof open | undefined;

  public get name(): string {
    return commands.OPEN;
  }

  public get description(): string {
    return 'Opens Azure AD app in the Azure AD portal';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.preview = typeof args.options.preview !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {    
    this.logOrOpenUrl(args, logger)
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private logOrOpenUrl(args: CommandArgs, logger: Logger): Promise<void> {
    return new Promise((resolve, reject) => {            
      const previewPrefix = args.options.preview === true ? "preview." : "";
      const url = `https://${previewPrefix}portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${this.appId}/isMSAApp/`;

      if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
        logger.log(`Use a web browser to open the page ${url}`);
        return resolve();
      }      
      
      logger.log(`Opening the following page in your browser: ${url}`);
      
      // 'open' is required here so we can lazy load the dependency.
      // _open is never set before hitting this line, but this check
      // is implemented so that we can stub it when testing.
      /* c8 ignore next 3 */
      if (!this._open) {
        this._open = require('open');
      }

      (this._open as typeof open)(url).then(() => {
        resolve();
      }, (error) => {
        reject(error);
      });      
    });
    
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--appId [appId]' },
      { option: '--preview' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {    
    if (args.options.appId && !validation.isValidGuid(args.options.appId as string)) {
      return `${args.options.appId} is not a valid GUID`;
    }
    
    return true;
  }
}

module.exports = new AppOpenCommand();