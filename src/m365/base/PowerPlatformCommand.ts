import auth, { CloudType } from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandArgs, CommandError } from '../../Command.js';
import { accessToken } from '../../utils/accessToken.js';


export default abstract class PowerPlatformCommand extends Command {
  protected get resource(): string {
    return 'https://api.bap.microsoft.com';
  }

  protected initAction(args: CommandArgs, logger: Logger): void {
    super.initAction(args, logger);

    if (!auth.connection.active) {
      // we fail no login in the base command command class
      return;
    }

    if (auth.connection.cloudType !== CloudType.Public) {
      throw new CommandError(`Power Platform commands only support the public cloud at the moment. We'll add support for other clouds in the future. Sorry for the inconvenience.`);
    }

    accessToken.assertDelegatedAccessToken();
  }
}
