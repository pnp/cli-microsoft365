import * as chalk from 'chalk';
import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.debug) {
      logger.log('Generating command completion...');
    }

    autocomplete.generateShCompletion();

    if (this.debug) {
      logger.log(chalk.green('DONE'));
    }

    cb();
  }
}

module.exports = new CliCompletionShUpdateCommand();