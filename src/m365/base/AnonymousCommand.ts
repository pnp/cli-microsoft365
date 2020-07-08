import Command, { CommandAction, CommandArgs } from '../../Command';

export default abstract class AnonymousCommand extends Command {
  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: (err?: any) => void) {
      args = (cmd as any).processArgs(args);
      (cmd as any).initAction(args, this);
      cmd.commandAction(this, args, cb);
    }
  }
}