import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';

class EntraUserRecycleBinItemListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Lists users from the recycle bin in the current tenant';
  }

  public alias(): string[] | undefined {
    return [aadCommands.USER_RECYCLEBINITEM_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'userPrincipalName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    this.showDeprecationWarning(logger, aadCommands.USER_RECYCLEBINITEM_LIST, commands.USER_RECYCLEBINITEM_LIST);

    if (this.verbose) {
      await logger.logToStderr('Retrieving users from the recycle bin...');
    }
    try {
      const users = await odata.getAllItems<User>(`${this.resource}/v1.0/directory/deletedItems/microsoft.graph.user`);
      await logger.log(users);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}
export default new EntraUserRecycleBinItemListCommand();