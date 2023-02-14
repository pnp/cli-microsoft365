import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

class AadUserRecycleBinItemListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Lists users from the recycle bin in the current tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'userPrincipalName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving users from the recycle bin...`);
    }
    try {
      const users = await odata.getAllItems<User>(`${this.resource}/v1.0/directory/deletedItems/microsoft.graph.user`);
      logger.log(users);
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }
}
module.exports = new AadUserRecycleBinItemListCommand();