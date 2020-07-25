import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandError
} from '../../../../Command';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import AnonymousCommand from '../../../base/AnonymousCommand';
import { autocomplete } from '../../../../autocomplete';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

    autocomplete.generateShCompletion();

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
        cmd.log(chalk.green('DONE'));
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
}

module.exports = new CliCompletionPwshSetupCommand();