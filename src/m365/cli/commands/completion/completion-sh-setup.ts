import * as chalk from 'chalk';
import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.debug) {
      logger.logToStderr('Generating command completion...');
    }

    autocomplete.generateShCompletion();

    if (this.debug) {
      logger.logToStderr('Registering command completion with the shell...');
    }

    autocomplete.setupShCompletion();

    logger.log('Command completion successfully registered. Restart your shell to load the completion');

    if (this.verbose) {
      logger.logToStderr(chalk.green('DONE'));
    }
    cb();
  }
}

module.exports = new CliCompletionShSetupCommand();