import * as fs from 'fs';
import { Cli } from '../../../cli/Cli';
import { Logger } from '../../../cli/Logger';
import { CommandError } from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import AnonymousCommand from '../../base/AnonymousCommand';
import { M365RcJson } from '../../base/M365RcJson';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
}

class ContextRemoveCommand extends AnonymousCommand {
  public get name(): string {
    return commands.REMOVE;
  }

  public get description(): string {
    return 'Removes the CLI for Microsoft 365 context in the current working folder';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.removeContext();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the context?`
      });

      if (result.continue) {
        await this.removeContext();
      }
    }
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--confirm'
      }
    );
  }

  private removeContext(): void {
    const filePath: string = '.m365rc.json';

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      try {
        const fileContents: string = fs.readFileSync(filePath, 'utf8');
        if (fileContents) {
          m365rc = JSON.parse(fileContents);
        }
      }
      catch (e) {
        throw new CommandError(`Error reading ${filePath}: ${e}. Please remove context info from ${filePath} manually.`);
      }
    }

    if (!m365rc.context) {
      return;
    }

    const keys = Object.keys(m365rc);
    if (keys.length === 1 && keys.indexOf('context') > -1) {
      try {
        fs.unlinkSync(filePath);
      }
      catch (e) {
        throw new CommandError(`Error removing ${filePath}: ${e}. Please remove ${filePath} manually.`);
      }
    }
    else {
      try {
        delete m365rc.context;
        fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
      }
      catch (e) {
        throw new CommandError(`Error writing ${filePath}: ${e}. Please remove context info from ${filePath} manually.`);
      }
    }
  }
}

module.exports = new ContextRemoveCommand();