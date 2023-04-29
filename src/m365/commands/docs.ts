import AnonymousCommand from '../base/AnonymousCommand';
import { Cli } from '../../cli/Cli';
import commands from './commands';
import { Logger } from '../../cli/Logger';
import type * as open from 'open';
import { settingsNames } from '../../settingsNames';
const packageJSON = require('../../../package.json');

class DocsCommand extends AnonymousCommand {
  private _open: typeof open | undefined;

  public get name(): string {
    return commands.DOCS;
  }

  public get description(): string {
    return 'Returns the CLI for Microsoft 365 docs webpage URL';
  }

  public async commandAction(logger: Logger): Promise<void> {
    logger.log(packageJSON.homepage);

    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      logger.log(`Use a web browser to open the CLI for Microsoft 365 docs webpage URL`);
      return;
    }

    // _open is never set before hitting this line, but this check
    // is implemented so that we can support lazy loading
    // but also stub it for testing
    /* c8 ignore next 3 */
    if (!this._open) {
      this._open = require('open');
    }
    await (this._open as typeof open)(packageJSON.homepage);
  }
}

module.exports = new DocsCommand();