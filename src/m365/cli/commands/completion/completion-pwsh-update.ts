import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class CliCompletionPwshUpdateCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_PWSH_UPDATE;
  }

  public get description(): string {
    return 'Updates command completion for PowerShell';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.debug) {
      logger.logToStderr('Generating command completion...');
    }

    autocomplete.generateShCompletion();

    cb();
  }
}

module.exports = new CliCompletionPwshUpdateCommand();