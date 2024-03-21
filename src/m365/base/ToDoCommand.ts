import { CommandArgs } from '../../Command.js';
import { Logger } from '../../cli/Logger.js';
import { accessToken } from '../../utils/accessToken.js';
import GraphCommand from './GraphCommand.js';

export default abstract class ToDoCommand extends GraphCommand {
  protected initAction(args: CommandArgs, logger: Logger): void {
    super.initAction(args, logger);

    accessToken.ensureDelegatedAccessToken();
  }
}