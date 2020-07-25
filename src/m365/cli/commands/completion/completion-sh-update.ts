import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import { autocomplete } from '../../../../autocomplete';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: GlobalOptions;
}

class CliCompletionShUpdateCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_SH_UPDATE;
  }

  public get description(): string {
    return 'Updates command completion for Zsh, Bash and Fish';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.debug) {
      cmd.log('Generating command completion...');
    }

    autocomplete.generateShCompletion();

    if (this.debug) {
      cmd.log(chalk.green('DONE'));
    }

    cb();
  }
}

module.exports = new CliCompletionShUpdateCommand();