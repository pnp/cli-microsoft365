import { Logger } from '../../../../cli';
import type * as open from 'open';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId: string;
  preview?: boolean;
  autoOpenBrowser?: boolean;
}

class AadAppOpenCommand extends GraphCommand {
  private _open: typeof open | undefined;

  public get name(): string {
    return commands.APP_OPEN;
  }

  public get description(): string {
    return 'Opens Azure AD app in the Azure AD portal';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.preview = typeof args.options.preview !== 'undefined';
    telemetryProps.autoOpenBrowser = typeof args.options.autoOpenBrowser !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .checkAppRegistrationExists(args, logger)
      .then(_ => this.logOrOpenUrl(args, logger))
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private checkAppRegistrationExists(args: CommandArgs, logger: Logger): Promise<void> {    
    const { appId } = args.options;

    if (this.verbose) {
      logger.logToStderr(`Check if Azure AD app ${appId} exists...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=appId eq '${encodeURIComponent(appId)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { id: string }[] }>(requestOptions)
      .then((res: { value: { id: string }[] }): Promise<void> => {
        if (res.value.length === 1) {
          return Promise.resolve();
        }

        return Promise.reject(`No Azure AD application registration with ID ${appId} found`);
      });
  }

  private logOrOpenUrl(args: CommandArgs, logger: Logger): Promise<void> {
    return new Promise((resolve, reject) => {
      const { appId } = args.options;
      const previewPrefix = args.options.preview === true ? "preview." : "";
      const url = `https://${previewPrefix}portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`;

      if (!args.options.autoOpenBrowser) {
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
      { option: '--appId <appId>' },
      { option: '--preview' },
      { option: '--autoOpenBrowser' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {    
    if (!validation.isValidGuid(args.options.appId as string)) {
      return `${args.options.appId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadAppOpenCommand();