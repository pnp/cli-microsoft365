import { Logger } from '../../../../cli';
import { CommandError, CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { GroupSettingTemplate } from './GroupSettingTemplate';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getAllItems(`${this.resource}/v1.0/groupSettingTemplates`, logger, true)
      .then((): void => {
        const groupSettingTemplate: GroupSettingTemplate[] = this.items.filter(t => args.options.id ? t.id === args.options.id : t.displayName === args.options.displayName);

        if (groupSettingTemplate && groupSettingTemplate.length > 0) {
          logger.log(groupSettingTemplate.pop());
        }
        else {
          cb(new CommandError(`Resource '${(args.options.id || args.options.displayName)}' does not exist.`));
          return;
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --displayName [displayName]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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
  }
}

module.exports = new AadGroupSettingTemplateGetCommand();