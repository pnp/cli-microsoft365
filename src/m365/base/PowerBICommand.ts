import Command, { CommandArgs } from '../../Command.js';
import auth from '../../Auth.js';
import { accessToken } from '../../utils/accessToken.js';
import { Logger } from '../../cli/Logger.js';

export default abstract class PowerBICommand extends Command {
  protected get resource(): string {
    return 'https://api.powerbi.com';
  }

  protected initAction(args: CommandArgs, logger: Logger): void {
    super.initAction(args, logger);

    if (!auth.connection.active) {
      // we fail no login in the base command command class
      return;
    }

    accessToken.ensureDelegatedAccessToken();
  }

}
