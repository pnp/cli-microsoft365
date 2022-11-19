import { Logger } from '../../../cli/Logger';
import ContextCommand from '../../base/ContextCommand';

import commands from '../commands';

class ContextInitCommand extends ContextCommand {
  public get name(): string {
    return commands.REMOVE;
  }

  public get description(): string {
    return 'Removes the CLI for Microsoft 365 context in the current working folder';
  }

  public async commandAction(logger: Logger): Promise<void> {
    await this.removeContextInfo(logger);
  }
}

module.exports = new ContextInitCommand();