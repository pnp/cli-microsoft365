import { Logger } from '../../cli/Logger.js';
import { app } from '../../utils/app.js';
import AnonymousCommand from '../base/AnonymousCommand.js';
import commands from './commands.js';

class VersionCommand extends AnonymousCommand {
  public get name(): string {
    return commands.VERSION;
  }

  public get description(): string {
    return 'Shows CLI for Microsoft 365 version';
  }

  public async commandAction(logger: Logger): Promise<void> {
    await logger.log(`v${app.packageJson().version}`);
  }
}

export default new VersionCommand();