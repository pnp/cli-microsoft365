import { Logger } from '../../cli';
import Command, { CommandArgs } from '../../Command';

export default abstract class AnonymousCommand extends Command {
  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    this.initAction(args, logger);
    await this.commandAction(logger, args);
  }
}