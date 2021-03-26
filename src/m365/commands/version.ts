import { Logger } from '../../cli';
import AnonymousCommand from '../base/AnonymousCommand';
import commands from './commands';
const packageJSON = require('../../../package.json');

class VersionCommand extends AnonymousCommand {
  public get name(): string {
    return commands.VERSION;
  }

  public get description(): string {
    return 'Shows CLI for Microsoft 365 version';
  }

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {
    logger.log(`v${packageJSON.version}`);
    cb();
  }
}

module.exports = new VersionCommand();