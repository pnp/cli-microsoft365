import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

class CliCompletionPwshUpdateCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_PWSH_UPDATE;
  }

  public get description(): string {
    return 'Updates command completion for PowerShell';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.debug) {
      logger.logToStderr('Generating command completion...');
    }

    autocomplete.generateShCompletion();
  }
}

module.exports = new CliCompletionPwshUpdateCommand();