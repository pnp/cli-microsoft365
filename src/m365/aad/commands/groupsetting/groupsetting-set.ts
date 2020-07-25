import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import { GroupSetting } from './GroupSetting';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class AadGroupSettingSetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.GROUPSETTING_SET}`;
  }

  public get description(): string {
    return 'Updates the particular group setting';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving group setting with id '${args.options.id}'...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    };

    request
      .get<GroupSetting>(requestOptions)
      .then((groupSetting: GroupSetting): Promise<{}> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          body: {
            displayName: groupSetting.displayName,
            templateId: groupSetting.templateId,
            values: this.getGroupSettingValues(args.options, groupSetting)
          },
          json: true
        };

        return request.patch(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getGroupSettingValues(options: any, groupSetting: GroupSetting): { name: string; value: string }[] {
    const values: { name: string; value: string }[] = [];
    const excludeOptions: string[] = [
      'id',
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

    groupSetting.values.forEach(v => {
      if (!values.find(e => e.name === v.name)) {
        values.push({
          name: v.name,
          value: v.value
        });
      }
    });

    return values;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the group setting to update'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new AadGroupSettingSetCommand();