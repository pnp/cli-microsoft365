
import * as fs from 'fs';
import { Logger } from '../../../cli/Logger';
import Command from '../../../Command';
import { M365RcJson } from '../../base/M365RcJson';
import commands from '../commands';

class ContextInitCommand extends Command {
  public get name(): string {
    return commands.INIT;
  }

  public get description(): string {
    return 'Retrieve context from Microsoft 365';
  }

  constructor() {
    super();
  }

  public async commandAction(logger: Logger): Promise<void> {
    await this.saveContextInfo({}, logger);
  }

  private saveContextInfo(contextInfo: any, logger: Logger): Promise<any> {
    const filePath: string = '.m365rc.json';

    if (this.verbose) {
      logger.logToStderr(`Saving context information to the ${filePath} file...`);
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
        logger.logToStderr(`Error reading ${filePath}: ${e}. Please add context info to ${filePath} manually.`);
        return contextInfo;
      }
    }

    if (!m365rc.context) {
      m365rc.context = contextInfo;

      try {
        fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
      }
      catch (e) {
        logger.logToStderr(`Error writing ${filePath}: ${e}. Please add context info to ${filePath} manually.`);
      }
    }

    return contextInfo;
  }

}

module.exports = new ContextInitCommand();