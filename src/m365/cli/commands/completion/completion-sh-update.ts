import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import { autocomplete } from '../../../../autocomplete';

const vorpal: Vorpal = require('../../../../vorpal-init');

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

    autocomplete.generateShCompletion(vorpal);

    if (this.debug) {
      cmd.log(vorpal.chalk.green('DONE'));
    }

    cb();
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(commands.COMPLETION_SH_UPDATE).helpInformation());
    log(
      `  Remarks:
  
    This commands updates the list of commands and their options used by
    command completion in Zsh, Bash and Fish. You should run this command each
    time after upgrading the CLI for Microsoft 365.
   
  Examples:
  
    Update list of commands for Zsh, Bash and Fish command completion
      ${this.getCommandName()}

  More information:

    Command completion
      https://pnp.github.io/cli-microsoft365/concepts/completion/
`);
  }
}

module.exports = new CliCompletionShUpdateCommand();