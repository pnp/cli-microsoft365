import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import { autocomplete } from '../../../../autocomplete';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    cmd.log(autocomplete.getClinkCompletion());
    cb();
  }
}

module.exports = new CliCompletionClinkUpdateCommand();