import type * as open from 'open';
import { Cli } from '../../../cli/Cli';
import { Logger } from '../../../cli/Logger';
import config from '../../../config';
import { settingsNames } from '../../../settingsNames';
import AnonymousCommand from '../../base/AnonymousCommand';
import commands from '../commands';

class CliReconsentCommand extends AnonymousCommand {
  private _open: typeof open | undefined;

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

    // _open is never set before hitting this line, but this check
    // is implemented so that we can support lazy loading
    // but also stub it for testing
    /* c8 ignore next 3 */
    if (!this._open) {
      this._open = require('open');
    }

    try {
      await (this._open as typeof open)(url);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new CliReconsentCommand();