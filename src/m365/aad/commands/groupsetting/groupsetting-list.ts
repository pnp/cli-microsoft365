import * as chalk from 'chalk';
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
    return `${commands.GROUPSETTING_LIST}`;
  }

  public get description(): string {
    return 'Lists Azure AD group settings';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/groupSettings`, logger, true)
      .then((): void => {
        if (args.options.output === 'json') {
          logger.log(this.items);
        }
        else {
          logger.log(this.items.map(i => {
            return {
              id: i.id,
              displayName: i.displayName
            };
          }));
        }

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadGroupSettingListCommand();