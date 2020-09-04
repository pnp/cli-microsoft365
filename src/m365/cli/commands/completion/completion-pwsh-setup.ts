import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import AnonymousCommand from '../../../base/AnonymousCommand';
import { autocomplete } from '../../../../autocomplete';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  profile: string;
}

class CliCompletionPwshSetupCommand extends AnonymousCommand {
  public get name(): string {
    return commands.COMPLETION_PWSH_SETUP;
  }

  public get description(): string {
    return 'Sets up command completion for PowerShell';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.debug) {
      cmd.log('Generating command completion...');
    }

    autocomplete.generateShCompletion(vorpal);

    if (this.debug) {
      cmd.log(`Ensuring that the specified profile path ${args.options.profile} exists...`);
    }

    if (fs.existsSync(args.options.profile)) {
      if (this.debug) {
        cmd.log('Profile file already exists');
      }
    }
    else {
      // check if the path exists
      const dirname: string = path.dirname(args.options.profile);
      if (fs.existsSync(dirname)) {
        if (this.debug) {
          cmd.log(`Profile path ${dirname} already exists`);
        }
      }
      else {
        try {
          if (this.debug) {
            cmd.log(`Profile path ${dirname} doesn't exist. Creating...`);
          }

          fs.mkdirSync(dirname, { recursive: true });
        }
        catch (e) {
          cb(new CommandError(e));
          return;
        }
      }

      if (this.debug) {
        cmd.log(`Creating profile file ${args.options.profile}...`);
      }

      try {
        fs.writeFileSync(args.options.profile, '', 'utf8');
      }
      catch (e) {
        cb(new CommandError(e));
        return;
      }
    }

    if (this.verbose) {
      cmd.log(`Adding CLI for Microsoft 365 command completion to PowerShell profile...`);
    }

    const completionScriptPath: string = path.resolve(__dirname, '..', '..', '..', '..', '..', 'scripts', 'Register-CLIM365Completion.ps1');
    try {
      fs.appendFileSync(args.options.profile, os.EOL + completionScriptPath, 'utf8');

      if (this.verbose) {
        cmd.log(vorpal.chalk.green('DONE'));
      }
      cb();
    }
    catch (e) {
      cb(new CommandError(e));
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --profile <profile>',
        description: 'Path to the PowerShell profile file'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.profile) {
        return 'Required option profile missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.COMPLETION_PWSH_SETUP).helpInformation());
    log(
      `  Remarks:
  
    This commands sets up command completion for the CLI for Microsoft 365 in
    PowerShell by registering a custom PowerShell argument completer
    in the specified profile. Because CLI for Microsoft 365 is not a native PowerShell
    module, it requires a custom completer to provide completion.
    
    If the specified profile path doesn't exist, the CLI will try to create it.
   
  Examples:
  
    Set up command completion for PowerShell using the profile from the ${chalk.grey('profile')}
    variable
      ${this.getCommandName()} --profile $profile

  More information:

    Command completion
      https://pnp.github.io/cli-microsoft365/concepts/completion/
`);
  }
}

module.exports = new CliCompletionPwshSetupCommand();