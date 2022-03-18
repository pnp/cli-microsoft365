import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { GroupSetting } from './GroupSetting';

interface CommandArgs {
  options: GlobalOptions;
}

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    odata
      .getAllItems<GroupSetting>(`${this.resource}/v1.0/groupSettings`, logger)
      .then((groupSettings): void => {
        logger.log(groupSettings);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadGroupSettingListCommand();