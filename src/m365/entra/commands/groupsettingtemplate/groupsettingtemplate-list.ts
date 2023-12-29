import { GroupSettingTemplate } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';

class AadGroupSettingTemplateListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTINGTEMPLATE_LIST;
  }

  public get description(): string {
    return 'Lists Entra ID group settings templates';
  }

  public alias(): string[] | undefined {
    return [aadCommands.GROUPSETTINGTEMPLATE_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const templates = await odata.getAllItems<GroupSettingTemplate>(`${this.resource}/v1.0/groupSettingTemplates`);
      await logger.log(templates);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadGroupSettingTemplateListCommand();