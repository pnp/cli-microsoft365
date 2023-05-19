import { Cli } from '../../../cli/Cli';
import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import { settingsNames } from '../../../settingsNames';
import { browserUtil } from '../../../utils/browserUtil';
import AppCommand from '../../base/AppCommand';
import commands from '../commands';

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

    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      logger.log(`Use a web browser to open the page ${url}`);
      return;
    }

    logger.log(`Opening the following page in your browser: ${url}`);
    await browserUtil.open(url);
  }
}

module.exports = new AppOpenCommand();