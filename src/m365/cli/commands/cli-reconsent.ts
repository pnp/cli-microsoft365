import { cli } from '../../../cli/cli.js';
import { Logger } from '../../../cli/Logger.js';
import { settingsNames } from '../../../settingsNames.js';
import { browserUtil } from '../../../utils/browserUtil.js';
import AnonymousCommand from '../../base/AnonymousCommand.js';
import commands from '../commands.js';

class CliReconsentCommand extends AnonymousCommand {
  public get name(): string {
    return commands.RECONSENT;
  }

  public get description(): string {
    return 'Returns URL to open in the browser to re-consent CLI for Microsoft 365 Microsoft Entra permissions';
  }

  public async commandAction(logger: Logger): Promise<void> {
    const url = `https://login.microsoftonline.com/${cli.getTenant()}/oauth2/authorize?client_id=${cli.getClientId()}&response_type=code&prompt=admin_consent`;

    if (cli.getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      await logger.log(`To re-consent your Microsoft Entra application navigate in your web browser to ${url}`);
      return;
    }

    await logger.log(`Opening the following page in your browser: ${url}`);

    try {
      await browserUtil.open(url);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new CliReconsentCommand();