import * as fs from 'fs';
import { Logger } from '../../../../cli/Logger';
import { CommandError } from '../../../../Command';
import ContextCommand from '../../../base/ContextCommand';
import { M365RcJson } from '../../../base/M365RcJson';
import commands from '../../commands';

class ContextOptionListCommand extends ContextCommand {
  public get name(): string {
    return commands.OPTION_LIST;
  }

  public get description(): string {
    return 'List all options added to the context';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving context options...`);
    }
    const filePath: string = '.m365rc.json';

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      try {
        if (this.verbose) {
          logger.logToStderr(`Reading context file...`);
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
      logger.log(m365rc.context);
    }
  }
}

module.exports = new ContextOptionListCommand();