import { Cli } from '../../../cli/Cli';
import { Logger } from '../../../cli/Logger';
import config from '../../../config';
import { settingsNames } from '../../../settingsNames';
import { browserUtil } from '../../../utils/browserUtil';
import AnonymousCommand from '../../base/AnonymousCommand';
import commands from '../commands';

class CliReconsentCommand extends AnonymousCommand {

  public get name(): string {
    return commands.RECONSENT;
  }

  public get description(): string {
    return 'Returns Azure AD URL to open in the browser to re-consent CLI for Microsoft 365 permissions';
  }

  public async commandAction(logger: Logger): Promise<void> {
    const url = `https://login.microsoftonline.com/${config.tenant}/oauth2/authorize?client_id=${config.cliAadAppId}&response_type=code&prompt=admin_consent`;

    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      logger.log(`To re-consent the PnP Microsoft 365 Management Shell Azure AD application navigate in your web browser to ${url}`);
      return;
    }

    logger.log(`Opening the following page in your browser: ${url}`);

    try {
      await browserUtil.open(url);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new CliReconsentCommand();