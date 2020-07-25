import Command, { CommandAction, CommandArgs } from '../../Command';
import { CommandInstance } from '../../cli';

export default abstract class AnonymousCommand extends Command {
  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: (err?: any) => void) {
      (cmd as any).initAction(args, this);
      cmd.commandAction(this, args, cb);
    }
  }
}