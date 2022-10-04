import { Logger } from '../../cli/Logger';
import AnonymousCommand from '../base/AnonymousCommand';
import commands from './commands';
const packageJSON = require('../../../package.json');

class VersionCommand extends AnonymousCommand {
  public get name(): string {
    return commands.VERSION;
  }

  public get description(): string {
    return 'Shows CLI for Microsoft 365 version';
  }

  public async commandAction(logger: Logger): Promise<void> {
    logger.log(`v${packageJSON.version}`);
  }
}

module.exports = new VersionCommand();