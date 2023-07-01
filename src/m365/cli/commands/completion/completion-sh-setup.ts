import { autocomplete } from '../../../../autocomplete.js';
import { Logger } from '../../../../cli/Logger.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';

class CliCompletionShSetupCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_SH_SETUP;
  }

  public get description(): string {
    return 'Sets up command completion for Zsh, Bash and Fish';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.debug) {
      await logger.logToStderr('Generating command completion...');
    }

    autocomplete.generateShCompletion();

    if (this.debug) {
      await logger.logToStderr('Registering command completion with the shell...');
    }

    autocomplete.setupShCompletion();

    await logger.log('Command completion successfully registered. Restart your shell to load the completion');
  }
}

export default new CliCompletionShSetupCommand();