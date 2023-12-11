import { cli } from '../../../cli/cli.js';
import { Logger } from '../../../cli/Logger.js';
import GlobalOptions from '../../../GlobalOptions.js';
import { settingsNames } from '../../../settingsNames.js';
import { browserUtil } from '../../../utils/browserUtil.js';
import AppCommand from '../../base/AppCommand.js';
import commands from '../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  preview?: boolean;
}

class AppOpenCommand extends AppCommand {
  public get name(): string {
    return commands.OPEN;
  }

  public get description(): string {
    return 'Opens Azure AD app in the Azure AD portal';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        appId: typeof args.options.appId !== 'undefined',
        preview: typeof args.options.preview !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' },
      { option: '--preview' }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await this.logOrOpenUrl(args, logger);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async logOrOpenUrl(args: CommandArgs, logger: Logger): Promise<void> {
    const previewPrefix = args.options.preview === true ? "preview." : "";
    const url = `https://${previewPrefix}portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${this.appId}/isMSAApp/`;

    if (cli.getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      await logger.log(`Use a web browser to open the page ${url}`);
      return;
    }

    await logger.log(`Opening the following page in your browser: ${url}`);
    await browserUtil.open(url);
  }
}

export default new AppOpenCommand();