import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import AnonymousCommand from '../../../base/AnonymousCommand';
import commands from '../../commands';

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

  constructor() {
    super();
  
    this.#initOptions();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --profile <profile>'
      }
    );
  }
  
  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.debug) {
      logger.logToStderr('Generating command completion...');
    }

    autocomplete.generateShCompletion();

    if (this.debug) {
      logger.logToStderr(`Ensuring that the specified profile path ${args.options.profile} exists...`);
    }

    if (fs.existsSync(args.options.profile)) {
      if (this.debug) {
        logger.logToStderr('Profile file already exists');
      }
    }
    else {
      // check if the path exists
      const dirname: string = path.dirname(args.options.profile);
      if (fs.existsSync(dirname)) {
        if (this.debug) {
          logger.logToStderr(`Profile path ${dirname} already exists`);
        }
      }
      else {
        try {
          if (this.debug) {
            logger.logToStderr(`Profile path ${dirname} doesn't exist. Creating...`);
          }

          fs.mkdirSync(dirname, { recursive: true });
        }
        catch (e: any) {
          cb(new CommandError(e));
          return;
        }
      }

      if (this.debug) {
        logger.logToStderr(`Creating profile file ${args.options.profile}...`);
      }

      try {
        fs.writeFileSync(args.options.profile, '', 'utf8');
      }
      catch (e: any) {
        cb(new CommandError(e));
        return;
      }
    }

    if (this.verbose) {
      logger.logToStderr(`Adding CLI for Microsoft 365 command completion to PowerShell profile...`);
    }

    const completionScriptPath: string = path.resolve(__dirname, '..', '..', '..', '..', '..', 'scripts', 'Register-CLIM365Completion.ps1');
    try {
      fs.appendFileSync(args.options.profile, os.EOL + completionScriptPath, 'utf8');
      cb();
    }
    catch (e: any) {
      cb(new CommandError(e));
    }
  }
}

module.exports = new CliCompletionPwshSetupCommand();