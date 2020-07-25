import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import { autocomplete } from '../../../../autocomplete';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: GlobalOptions;
}

class CliCompletionShSetupCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_SH_SETUP;
  }

  public get description(): string {
    return 'Sets up command completion for Zsh, Bash and Fish';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.debug) {
      cmd.log('Generating command completion...');
    }

    autocomplete.generateShCompletion();

    if (this.debug) {
      cmd.log('Registering command completion with the shell...');
    }

    autocomplete.setupShCompletion();

    cmd.log('Command completion successfully registered. Restart your shell to load the completion');

    if (this.verbose) {
      cmd.log(chalk.green('DONE'));
    }
    cb();
  }
}

module.exports = new CliCompletionShSetupCommand();