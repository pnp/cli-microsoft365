import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

class CliCompletionShSetupCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_SH_SETUP;
  }

  public get description(): string {
    return 'Sets up command completion for Zsh, Bash and Fish';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.debug) {
      logger.logToStderr('Generating command completion...');
    }

    autocomplete.generateShCompletion();

    if (this.debug) {
      logger.logToStderr('Registering command completion with the shell...');
    }

    autocomplete.setupShCompletion();

    logger.log('Command completion successfully registered. Restart your shell to load the completion');
  }
}

module.exports = new CliCompletionShSetupCommand();