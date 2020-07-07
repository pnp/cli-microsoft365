import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import { GroupSettingTemplate } from '../groupsettingtemplate/GroupSettingTemplate';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  templateId: string;
}

class AadGroupSettingAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.GROUPSETTING_ADD}`;
  }

  public get description(): string {
    return 'Creates a group setting';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.templateId = args.options.templateId;
    return telemetryProps;
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving group setting template with id '${args.options.templateId}'...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groupSettingTemplates/${args.options.templateId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    };

    request
      .get<GroupSettingTemplate>(requestOptions)
      .then((groupSettingTemplate: GroupSettingTemplate): Promise<{}> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/groupSettings`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          body: {
            templateId: args.options.templateId,
            values: this.getGroupSettingValues(args.options, groupSettingTemplate)
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getGroupSettingValues(options: any, groupSettingTemplate: GroupSettingTemplate): { name: string; value: string }[] {
    const values: { name: string; value: string }[] = [];
    const excludeOptions: string[] = [
      'templateId',
      'debug',
      'verbose',
      'output'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        values.push({
          name: key,
          value: options[key]
        });
      }
    });

    groupSettingTemplate.values.forEach(v => {
      if (!values.find(e => e.name === v.name)) {
        values.push({
          name: v.name,
          value: v.defaultValue
        });
      }
    });

    return values;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --templateId <templateId>',
        description: 'The ID of the group setting template to use to create the group setting'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.templateId) {
        return 'Required option templateId missing';
      }

      if (!Utils.isValidGuid(args.options.templateId)) {
        return `${args.options.templateId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    To create a group setting, you have to specify the ID of the group setting
    template that should be used to create the setting. You can retrieve the ID
    of the template using the ${chalk.blue(commands.GROUPSETTINGTEMPLATE_LIST)}
    command.

    To specify values for the different properties specified in the group
    setting template, include additional options that match the property in the
    group setting template.
    For example ${chalk.blue("--ClassificationList 'HBI, MBI, LBI, GDPR'")} will set
    the list of classifications to use on modern SharePoint sites.

    Each group setting template specifies default value for each property. If
    you don't specify a value for the particular property yourself, the default
    value from the group setting template will be used. To find out which
    properties are available for the particular group setting template, use the
    ${chalk.blue(commands.GROUPSETTINGTEMPLATE_GET)} command.

    If the specified ${chalk.blue('templateId')} doesn't reference a valid group setting
    template, you will get a ${chalk.grey("Resource 'xyz' does not exist or one of its ")}
    ${chalk.grey('queried reference-property objects are not present.')} error.

    If you try to add a group setting using a template, for which a setting
    already exists, you will get a ${chalk.grey('A conflicting object with one or more ')}
    ${chalk.grey('of the specified property values is present in the directory.')} error.

  Examples:
  
    Configure classification for modern SharePoint sites
      ${this.name} --templateId 62375ab9-6b52-47ed-826b-58e47e0e304b --UsageGuidelinesUrl https://contoso.sharepoint.com/sites/compliance --ClassificationList 'HBI, MBI, LBI, GDPR' --DefaultClassification MBI
`);
  }
}

module.exports = new AadGroupSettingAddCommand();