import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli/Logger';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

class CliCompletionClinkUpdateCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_CLINK_UPDATE;
  }

  public get description(): string {
    return 'Updates command completion for Clink (cmder)';
  }

  public async commandAction(logger: Logger): Promise<void> {
    logger.log(autocomplete.getClinkCompletion());
  }
}

module.exports = new CliCompletionClinkUpdateCommand();