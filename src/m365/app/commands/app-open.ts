import { z } from 'zod';
import { cli } from '../../../cli/cli.js';
import { Logger } from '../../../cli/Logger.js';
import { settingsNames } from '../../../settingsNames.js';
import { browserUtil } from '../../../utils/browserUtil.js';
import AppCommand, { appCommandOptions } from '../../base/AppCommand.js';
import commands from '../commands.js';

export const options = z.strictObject({
  ...appCommandOptions.shape,
  preview: z.boolean().optional().default(false)
});
type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class AppOpenCommand extends AppCommand {
  public get name(): string {
    return commands.OPEN;
  }

  public get description(): string {
    return 'Opens Microsoft Entra app in the Microsoft Entra ID portal';
  }

  public get schema(): z.ZodType | undefined {
    return options;
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