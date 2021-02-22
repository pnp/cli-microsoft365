import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { GroupSettingTemplate } from './GroupSettingTemplate';

interface CommandArgs {
  options: GlobalOptions;
}

class AadGroupSettingTemplateListCommand extends GraphItemsListCommand<GroupSettingTemplate> {
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
    this
      .getAllItems(`${this.resource}/v1.0/groupSettingTemplates`, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadGroupSettingTemplateListCommand();