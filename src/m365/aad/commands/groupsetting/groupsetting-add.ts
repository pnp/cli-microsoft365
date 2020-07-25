import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import { GroupSettingTemplate } from '../groupsettingtemplate/GroupSettingTemplate';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
          cmd.log(chalk.green('DONE'));
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
      if (!Utils.isValidGuid(args.options.templateId)) {
        return `${args.options.templateId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new AadGroupSettingAddCommand();