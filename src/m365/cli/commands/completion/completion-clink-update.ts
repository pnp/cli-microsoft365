import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import { autocomplete } from '../../../../autocomplete';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    cmd.log(autocomplete.getClinkCompletion(vorpal));
    cb();
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.COMPLETION_CLINK_UPDATE).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} Before running this command, change the working directory
    to where your shell stores completion plugins. For cmder, it's
    ${chalk.grey('%CMDER_ROOT%\\vendor\\clink-completions')}, where ${chalk.grey('%CMDER_ROOT%')} is the folder
    where you installed cmder. After running this command, restart your terminal
    to load the completion.
      
  Remarks:
  
    This commands updates the list of commands and their options used by
    command completion in Clink (cmder). You should run this command each time
    after upgrading the CLI for Microsoft 365.
   
  Examples:
  
    Write the list of commands for Clink (cmder) command completion to a file
    named ${chalk.grey('m365.lua')} in the current directory
      ${this.getCommandName()} > m365.lua

  More information:
    Command completion
      https://pnp.github.io/cli-microsoft365/concepts/completion/
`);
  }
}

module.exports = new CliCompletionClinkUpdateCommand();