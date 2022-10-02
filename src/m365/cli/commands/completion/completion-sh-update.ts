import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli/Logger';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

class CliCompletionShUpdateCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_SH_UPDATE;
  }

  public get description(): string {
    return 'Updates command completion for Zsh, Bash and Fish';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.debug) {
      logger.logToStderr('Generating command completion...');
    }

    autocomplete.generateShCompletion();
  }
}

module.exports = new CliCompletionShUpdateCommand();