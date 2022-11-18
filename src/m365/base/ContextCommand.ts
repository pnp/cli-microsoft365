import * as fs from 'fs';
import AnonymousCommand from './AnonymousCommand';
import { Logger } from '../../cli/Logger';
import { M365RcJson } from './M365RcJson';
import { Hash } from '../../utils/types';

export default abstract class ContextCommand extends AnonymousCommand {
  public saveContextInfo(context: Hash, logger: Logger): Promise<any> {
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
        return Promise.resolve(context);
      }
    }

    if (!m365rc.context) {
      m365rc.context = context;

      try {
        fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
      }
      catch (e) {
        logger.logToStderr(`Error writing ${filePath}: ${e}. Please add context info to ${filePath} manually.`);
      }
    }

    return Promise.resolve(context);
  }
}
