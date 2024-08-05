import auth, { CloudType } from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandArgs, CommandError } from '../../Command.js';
import { accessToken } from '../../utils/accessToken.js';

export default abstract class PowerAppsCommand extends Command {
  protected get resource(): string {
    return 'https://api.powerapps.com';
  }

  protected async initAction(args: CommandArgs, logger: Logger): Promise<void> {
    await super.initAction(args, logger);

    if (!auth.connection.active) {
      // we fail no login in the base command command class
      return;
    }

    if (auth.connection.cloudType !== CloudType.Public) {
      throw new CommandError(`Power Apps commands only support the public cloud at the moment. We'll add support for other clouds in the future. Sorry for the inconvenience.`);
    }

    accessToken.assertDelegatedAccessToken();
  }
}