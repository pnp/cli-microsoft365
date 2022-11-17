import * as fs from 'fs';
import AnonymousCommand from './AnonymousCommand';
import { Logger } from '../../cli/Logger';
import { M365RcJson } from './M365RcJson';

export default abstract class ContextCommand extends AnonymousCommand {
  public saveContextInfo(contextInfo: any, logger: Logger): Promise<any> {
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
