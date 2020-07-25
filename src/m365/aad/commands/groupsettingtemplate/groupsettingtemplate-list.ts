import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { GroupSettingTemplate } from './GroupSettingTemplate';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: GlobalOptions;
}

class AadGroupSettingTemplateListCommand extends GraphItemsListCommand<GroupSettingTemplate> {
  public get name(): string {
    return `${commands.GROUPSETTINGTEMPLATE_LIST}`;
  }

  public get description(): string {
    return 'Lists Azure AD group settings templates';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/groupSettingTemplates`, cmd, true)
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

module.exports = new AadGroupSettingTemplateListCommand();