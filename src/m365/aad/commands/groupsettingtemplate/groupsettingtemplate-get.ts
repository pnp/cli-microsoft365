import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { GroupSettingTemplate } from './GroupSettingTemplate';
import { CommandError } from '../../../../Command';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
}

class AadGroupSettingTemplateGetCommand extends GraphItemsListCommand<GroupSettingTemplate> {
  public get name(): string {
    return `${commands.GROUPSETTINGTEMPLATE_GET}`;
  }

  public get description(): string {
    return 'Gets information about the specified Azure AD group settings template';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.displayName = typeof args.options.displayName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/groupSettingTemplates`, cmd, true)
      .then((): void => {
        const groupSettingTemplate: GroupSettingTemplate[] = this.items.filter(t => args.options.id ? t.id === args.options.id : t.displayName === args.options.displayName);

        if (groupSettingTemplate && groupSettingTemplate.length > 0) {
          cmd.log(groupSettingTemplate.pop());
        }
        else {
          cb(new CommandError(`Resource '${(args.options.id || args.options.displayName)}' does not exist.`));
          return;
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'The ID of the settings template to retrieve. Specify the id or displayName but not both'
      },
      {
        option: '-n, --displayName [displayName]',
        description: 'The display name of the settings template to retrieve. Specify the id or displayName but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id && !args.options.displayName) {
        return 'Specify either id or displayName';
      }

      if (args.options.id && args.options.displayName) {
        return 'Specify either id or displayName but not both';
      }

      if (args.options.id &&
        !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new AadGroupSettingTemplateGetCommand();