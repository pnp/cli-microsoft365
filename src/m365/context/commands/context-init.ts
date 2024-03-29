import ContextCommand from '../../base/ContextCommand.js';
import commands from '../commands.js';

class ContextInitCommand extends ContextCommand {
  public get name(): string {
    return commands.INIT;
  }

  public get description(): string {
    return 'Initiates CLI for Microsoft 365 context in the current working folder';
  }

  public async commandAction(): Promise<void> {
    await this.saveContextInfo({});
  }
}

export default new ContextInitCommand();