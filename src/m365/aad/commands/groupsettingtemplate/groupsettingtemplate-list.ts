import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { GroupSettingTemplate } from './GroupSettingTemplate';

interface CommandArgs {
  options: GlobalOptions;
}

class AadGroupSettingTemplateListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTINGTEMPLATE_LIST;
  }

  public get description(): string {
    return 'Lists Azure AD group settings templates';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    odata
      .getAllItems<GroupSettingTemplate>(`${this.resource}/v1.0/groupSettingTemplates`, logger)
      .then((templates): void => {
        logger.log(templates);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadGroupSettingTemplateListCommand();