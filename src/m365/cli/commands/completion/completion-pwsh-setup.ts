import fs from 'fs';
import os from 'os';
import path from 'path';
import url from 'url';
import { autocomplete } from '../../../../autocomplete.js';
import { Logger } from '../../../../cli/Logger.js';
import {
  CommandError
} from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.debug) {
      await logger.logToStderr('Generating command completion...');
    }

    autocomplete.generateShCompletion();

    if (this.debug) {
      await logger.logToStderr(`Ensuring that the specified profile path ${args.options.profile} exists...`);
    }

    if (fs.existsSync(args.options.profile)) {
      if (this.debug) {
        await logger.logToStderr('Profile file already exists');
      }
    }
    else {
      // check if the path exists
      const dirname: string = path.dirname(args.options.profile);
      if (fs.existsSync(dirname)) {
        if (this.debug) {
          await logger.logToStderr(`Profile path ${dirname} already exists`);
        }
      }
      else {
        try {
          if (this.debug) {
            await logger.logToStderr(`Profile path ${dirname} doesn't exist. Creating...`);
          }

          fs.mkdirSync(dirname, { recursive: true });
        }
        catch (e: any) {
          throw new CommandError(e);
        }
      }

      if (this.debug) {
        await logger.logToStderr(`Creating profile file ${args.options.profile}...`);
      }

      try {
        fs.writeFileSync(args.options.profile, '', 'utf8');
      }
      catch (e: any) {
        throw new CommandError(e);
      }
    }

    if (this.verbose) {
      await logger.logToStderr(`Adding CLI for Microsoft 365 command completion to PowerShell profile...`);
    }

    const completionScriptPath: string = path.resolve(__dirname, '..', '..', '..', '..', '..', 'scripts', 'Register-CLIM365Completion.ps1');
    try {
      fs.appendFileSync(args.options.profile, os.EOL + completionScriptPath, 'utf8');
      return;
    }
    catch (e: any) {
      throw new CommandError(e);
    }
  }
}

export default new CliCompletionPwshSetupCommand();