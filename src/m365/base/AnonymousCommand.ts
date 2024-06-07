import { Logger } from '../../cli/Logger.js';
import Command, { CommandArgs } from '../../Command.js';

export default abstract class AnonymousCommand extends Command {
  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    await this.initAction(args, logger);
    await this.commandAction(logger, args);
  }
}