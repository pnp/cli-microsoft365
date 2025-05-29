import auth from '../../Auth.js';
import { CommandArgs } from '../../Command.js';
import { Logger } from '../../cli/Logger.js';
import { accessToken } from '../../utils/accessToken.js';
import GraphCommand from './GraphCommand.js';

/**
 * This command class is for application-only Graph commands.  
 */
export default abstract class GraphApplicationCommand extends GraphCommand {
  protected async initAction(args: CommandArgs, logger: Logger): Promise<void> {
    await super.initAction(args, logger);

    if (!auth.connection.active) {
      // we fail no login in the base command command class
      return;
    }

    accessToken.assertAccessTokenType('application');
  }
}