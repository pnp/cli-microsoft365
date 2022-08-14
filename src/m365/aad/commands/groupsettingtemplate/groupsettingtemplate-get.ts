import { Logger } from '../../../../cli';
import { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { GroupSettingTemplate } from './GroupSettingTemplate';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
}

class AadGroupSettingTemplateGetCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTINGTEMPLATE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Azure AD group settings template';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --displayName [displayName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id &&
          !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'displayName']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    odata
      .getAllItems<GroupSettingTemplate>(`${this.resource}/v1.0/groupSettingTemplates`)
      .then((templates): void => {
        const groupSettingTemplate: GroupSettingTemplate[] = templates.filter(t => args.options.id ? t.id === args.options.id : t.displayName === args.options.displayName);

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
}

module.exports = new AadGroupSettingTemplateGetCommand();