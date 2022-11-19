import * as fs from 'fs';
import { Cli } from '../../../cli/Cli';
import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import ContextCommand from '../../base/ContextCommand';
import { M365RcJson } from '../../base/M365RcJson';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
}

class ContextRemoveCommand extends ContextCommand {
  public get name(): string {
    return commands.REMOVE;
  }

  public get description(): string {
    return 'Removes the CLI for Microsoft 365 context in the current working folder';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.removeContextInfo(logger);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Microsoft Power App ${args.options.name}?`
      });

      if (result.continue) {
        await this.removeContextInfo(logger);
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

  public removeContextInfo(logger: Logger): void {
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
        logger.logToStderr(`Error reading ${filePath}: ${e}. Please remove context info from ${filePath} manually.`);
        return;
      }
    }

    if (m365rc.context) {
      const keys = Object.keys(m365rc);
      if (keys.length === 1 && keys.indexOf('context') > -1) {
        try {
          fs.unlinkSync(filePath);
        }
        catch (e) {
          logger.logToStderr(`Error removing ${filePath}: ${e}. Please remove ${filePath} manually.`);
        }
      }
      else {
        try {
          delete m365rc.context;
          fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
        }
        catch (e) {
          logger.logToStderr(`Error writing ${filePath}: ${e}. Please remove context info from ${filePath} manually.`);
        }
      }
    }
  }
}

module.exports = new ContextRemoveCommand();