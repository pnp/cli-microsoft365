import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { GroupSetting } from './GroupSetting';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/groupSettings`, cmd, true)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map(i => {
            return {
              id: i.id,
              displayName: i.displayName
            };
          }));
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }
}

module.exports = new AadGroupSettingListCommand();