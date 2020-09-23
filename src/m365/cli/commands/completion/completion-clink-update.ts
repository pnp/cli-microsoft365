import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class CliCompletionClinkUpdateCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_CLINK_UPDATE;
  }

  public get description(): string {
    return 'Updates command completion for Clink (cmder)';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    logger.log(autocomplete.getClinkCompletion());
    cb();
  }
}

module.exports = new CliCompletionClinkUpdateCommand();