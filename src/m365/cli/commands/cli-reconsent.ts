import type * as open from 'open';
import { Cli, Logger } from '../../../cli';
import config from '../../../config';
import GlobalOptions from '../../../GlobalOptions';
import { settingsNames } from '../../../settingsNames';
import AnonymousCommand from '../../base/AnonymousCommand';
import commands from '../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class CliReconsentCommand extends AnonymousCommand {
  private _open: typeof open | undefined;

  public get name(): string {
    return commands.RECONSENT;
  }

  public get description(): string {
    return 'Returns Azure AD URL to open in the browser to re-consent CLI for Microsoft 365 permissions';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const url = `https://login.microsoftonline.com/${config.tenant}/oauth2/authorize?client_id=${config.cliAadAppId}&response_type=code&prompt=admin_consent`;

    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      logger.log(`To re-consent the PnP Microsoft 365 Management Shell Azure AD application navigate in your web browser to ${url}`);
      return cb();
    }

    logger.log(`Opening the following page in your browser: ${url}`);

    // _open is never set before hitting this line, but this check
    // is implemented so that we can support lazy loading
    // but also stub it for testing
    /* c8 ignore next 3 */
    if (!this._open) {
      this._open = require('open');
    }

    (this._open as typeof open)(url).then(() => {
      cb();
    }, (error) => {
      this.handleRejectedODataJsonPromise(error, logger, cb);
    });
  }
}

module.exports = new CliReconsentCommand();