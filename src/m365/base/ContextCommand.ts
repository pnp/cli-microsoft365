import * as fs from 'fs';
import AnonymousCommand from './AnonymousCommand';
import { Logger } from '../../cli/Logger';
import { M365RcJson } from './M365RcJson';

export default abstract class ContextCommand extends AnonymousCommand {
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
        logger.logToStderr(`Error reading ${filePath}: ${e}. Please remove context info to ${filePath} manually.`);
        return;
      }
    }

    if (m365rc.context) {
      const keys = Object.keys(m365rc);
      if (keys.length === 1 && keys.indexOf('context') > -1) {
        fs.unlinkSync(filePath);
      }
      else {
        delete m365rc.context;
        try {
          fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
        }
        catch (e) {
          logger.logToStderr(`Error writing ${filePath}: ${e}. Please remove context info to ${filePath} manually.`);
        }
      }
    }
  }
}