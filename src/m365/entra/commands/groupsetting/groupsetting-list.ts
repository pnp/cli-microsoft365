import { GroupSetting } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

class EntraGroupSettingListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTING_LIST;
  }

  public get description(): string {
    return 'Lists Entra group settings';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const groupSettings = await odata.getAllItems<GroupSetting>(`${this.resource}/v1.0/groupSettings`);
      await logger.log(groupSettings);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraGroupSettingListCommand();