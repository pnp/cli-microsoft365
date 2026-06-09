import { GroupSetting } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({ ...globalOptionsZod.shape });

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

  public get schema(): z.ZodType | undefined {
    return options;
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