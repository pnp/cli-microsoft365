import commands from './commands';
import AnonymousCommand from '../base/AnonymousCommand';
import { CommandInstance } from '../../cli';
const packageJSON = require('../../../package.json');

class VersionCommand extends AnonymousCommand {
  public get name(): string {
    return commands.VERSION;
  }

  public get description(): string {
    return 'Shows CLI for Microsoft 365 version';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: (err?: any) => void): void {
    cmd.log(`v${packageJSON.version}`);
    cb();
  }
}

module.exports = new VersionCommand();