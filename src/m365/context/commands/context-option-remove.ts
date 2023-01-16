import * as fs from 'fs';
import { Cli } from '../../../cli/Cli';
import { Logger } from '../../../cli/Logger';
import { CommandError } from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import ContextCommand from '../../base/ContextCommand';
import { M365RcJson } from '../../base/M365RcJson';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  confirm?: boolean;
}

class ContextOptionRemoveCommand extends ContextCommand {
  public get name(): string {
    return commands.OPTION_REMOVE;
  }

  public get description(): string {
    return 'Removes a CLI for Microsoft 365 context option in the current working folder';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initTelemetry();
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
        option: '-n, --name <name>'
      },
      {
        option: '--confirm'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      this.removeContextOption(args.options.name);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the context option ${args.options.name}?`
      });

      if (result.continue) {
        this.removeContextOption(args.options.name);
      }
    }
  }
  private removeContextOption(name: string): void {
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
        throw new CommandError(`Error reading ${filePath}: ${e}. Please remove context option ${name} from ${filePath} manually.`);
      }
    }

    if (!m365rc.context || !m365rc.context[name]) {
      throw new CommandError(`There is no option ${name} in the context info`);
    }
    else {
      try {
        delete m365rc.context[name];
        fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
      }
      catch (e) {
        throw new CommandError(`Error writing ${filePath}: ${e}. Please remove context option ${name} from ${filePath} manually.`);
      }
    }
  }
}

module.exports = new ContextOptionRemoveCommand();