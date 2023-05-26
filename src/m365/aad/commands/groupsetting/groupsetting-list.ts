import { GroupSetting } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

class AadGroupSettingListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTING_LIST;
  }

  public get description(): string {
    return 'Lists Azure AD group settings';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const groupSettings = await odata.getAllItems<GroupSetting>(`${this.resource}/v1.0/groupSettings`);
      logger.log(groupSettings);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadGroupSettingListCommand();