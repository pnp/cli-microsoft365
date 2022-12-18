import * as fs from 'fs';
import { CommandError } from '../../Command';
import { Hash } from '../../utils/types';
import AnonymousCommand from './AnonymousCommand';
import { M365RcJson } from './M365RcJson';

export default abstract class ContextCommand extends AnonymousCommand {
  protected saveContextInfo(context: Hash): void {
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
        throw new CommandError(`Error reading ${filePath}: ${e}. Please add context info to ${filePath} manually.`);
      }
    }

    if (!m365rc.context) {
      m365rc.context = context;

      try {
        fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
      }
      catch (e) {
        throw new CommandError(`Error writing ${filePath}: ${e}. Please add context info to ${filePath} manually.`);
      }
    }
  }
}