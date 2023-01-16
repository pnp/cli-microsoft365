import * as fs from 'fs';
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
  value: string;
}

class ContextOptionAddCommand extends ContextCommand {
  public get name(): string {
    return commands.OPTION_ADD;
  }

  public get description(): string {
    return 'Adds a CLI for Microsoft 365 context option in the current working folder';
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-v, --value <value>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const filePath: string = '.m365rc.json';

    if (this.verbose) {
      logger.logToStderr(`Saving ${args.options.name} with value ${args.options.value} to the ${filePath} file...`);
    }

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      if (this.debug) {
        logger.logToStderr(`Reading existing ${filePath}...`);
      }

      try {
        const fileContents: string = fs.readFileSync(filePath, 'utf8');
        if (fileContents) {
          m365rc = JSON.parse(fileContents);
        }
      }
      catch (e) {
        throw new CommandError(`Error reading ${filePath}: ${e}. Please add ${args.options.name} to ${filePath} manually.`);
      }
    }

    if (m365rc.context) {
      m365rc.context[args.options.name] = args.options.value;
      try {
        fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
      }
      catch (e) {
        throw new CommandError(`Error writing ${filePath}: ${e}. Please add ${args.options.name} to ${filePath} manually.`);
      }
    }
    else {
      this.saveContextInfo({ [args.options.name]: args.options.value });
    }
  }
}

module.exports = new ContextOptionAddCommand();