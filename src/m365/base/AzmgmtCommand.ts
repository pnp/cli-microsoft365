import auth, { CloudType } from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandArgs, CommandError } from '../../Command.js';

export default abstract class AzmgmtCommand extends Command {
  protected get resource(): string {
    return 'https://management.azure.com/';
  }

  protected initAction(args: CommandArgs, logger: Logger): void {
    super.initAction(args, logger);

    if (!auth.service.active) {
      // we fail no login in the base command command class
      return;
    }

    if (auth.service.cloudType !== CloudType.Public) {
      throw new CommandError(`Power Automate commands only support the public cloud at the moment. We'll add support for other clouds in the future. Sorry for the inconvenience.`);
    }
  }
}