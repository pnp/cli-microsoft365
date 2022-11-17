

import { Logger } from '../../../cli/Logger';
import ContextCommand from '../../base/ContextCommand';

import commands from '../commands';

class ContextInitCommand extends ContextCommand {
  public get name(): string {
    return commands.INIT;
  }

  public get description(): string {
    return 'Initiates CLI for Microsoft 365 context in the current working folder';
  }

  public async commandAction(logger: Logger): Promise<void> {
    const contextInfo = await this.saveContextInfo({}, logger);
    logger.log(contextInfo);
  }
}

module.exports = new ContextInitCommand();