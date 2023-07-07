import AnonymousCommand from '../base/AnonymousCommand';
import { Cli } from '../../cli/Cli';
import commands from './commands';
import { Logger } from '../../cli/Logger';
import { settingsNames } from '../../settingsNames';
import { browserUtil } from '../../utils/browserUtil';
const packageJSON = require('../../../package.json');

class DocsCommand extends AnonymousCommand {

  public get name(): string {
    return commands.DOCS;
  }

  public get description(): string {
    return 'Returns the CLI for Microsoft 365 docs webpage URL';
  }

  public async commandAction(logger: Logger): Promise<void> {
    logger.log(packageJSON.homepage);

    if (Cli.getInstance().getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      return;
    }

    await browserUtil.open(packageJSON.homepage);
  }
}

module.exports = new DocsCommand();