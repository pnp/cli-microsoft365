import { Logger } from '../../cli';
import Command, { CommandArgs } from '../../Command';

export default abstract class AnonymousCommand extends Command {
  public action(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this.initAction(args, logger);
    this.commandAction(logger, args, cb);
  }
}