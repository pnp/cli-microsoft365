import { autocomplete } from '../../../../autocomplete.js';
import { Logger } from '../../../../cli/Logger.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';

class CliCompletionClinkUpdateCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_CLINK_UPDATE;
  }

  public get description(): string {
    return 'Updates command completion for Clink (cmder)';
  }

  public async commandAction(logger: Logger): Promise<void> {
    await logger.log(autocomplete.getClinkCompletion());
  }
}

export default new CliCompletionClinkUpdateCommand();