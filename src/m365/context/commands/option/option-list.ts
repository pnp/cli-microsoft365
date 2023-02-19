import fs from 'fs';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import ContextCommand from '../../../base/ContextCommand.js';
import { M365RcJson } from '../../../base/M365RcJson.js';
import commands from '../../commands.js';

class ContextOptionListCommand extends ContextCommand {
  public get name(): string {
    return commands.OPTION_LIST;
  }

  public get description(): string {
    return 'List all options added to the context';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving context options...`);
    }
    const filePath: string = '.m365rc.json';

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Reading context file...`);
        }
        const fileContents: string = fs.readFileSync(filePath, 'utf8');
        if (fileContents) {
          m365rc = JSON.parse(fileContents);
        }
      }
      catch (e) {
        throw new CommandError(`Error reading ${filePath}: ${e}. Please retrieve context options from ${filePath} manually.`);
      }
    }

    if (!m365rc.context) {
      throw new CommandError(`No context present`);
    }
    else {
      await logger.log(m365rc.context);
    }
  }
}

export default new ContextOptionListCommand();