import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { GroupSetting } from './GroupSetting';

interface CommandArgs {
  options: GlobalOptions;
}

class AadGroupSettingListCommand extends GraphItemsListCommand<GroupSetting> {
  public get name(): string {
    return commands.GROUPSETTING_LIST;
  }

  public get description(): string {
    return 'Lists Azure AD group settings';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/groupSettings`, logger, true)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadGroupSettingListCommand();