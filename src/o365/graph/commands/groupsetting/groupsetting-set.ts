import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../GraphCommand';
import { GroupSetting } from './GroupSetting';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class GraphGroupSettingSetCommand extends GraphCommand {
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
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        if (this.verbose) {
          cmd.log(`Retrieving group setting with id '${args.options.id}'...`);
        }

        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/groupSettings/${args.options.id}`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((groupSetting: GroupSetting): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(groupSetting);
          cmd.log('');
        }

        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/groupSettings/${args.options.id}`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          }),
          body: {
            displayName: groupSetting.displayName,
            templateId: groupSetting.templateId,
            values: this.getGroupSettingValues(args.options, groupSetting)
          },
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.patch(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.id) {
        return 'Required option id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Microsoft Graph
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To update a group setting, you have to first connect to the Microsoft Graph
    using the ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT}`)}.

    To update a group setting, you have to specify the ID of the group setting.
    You can retrieve the ID of the group setting using the
    ${chalk.blue(commands.GROUPSETTING_LIST)} command.

    To update values for the different properties specified in the group
    setting, include additional options that match the property in the group
    setting. For example ${chalk.blue("--ClassificationList 'HBI, MBI, LBI, GDPR'")} will set
    the list of classifications to use on modern SharePoint sites.

    If you don't specify a value for the particular property, it will remain
    unchanged. To find out which properties are available for the particular
    group setting, use the ${chalk.blue(commands.GROUPSETTING_GET)} command.

    If the specified ${chalk.blue('id')} doesn't reference a valid group setting, you will get
    a ${chalk.grey("Resource 'xyz' does not exist or one of its queried reference-property")}
    ${chalk.grey('objects are not present.')} error.

  Examples:
  
    Configure classification for modern SharePoint sites
      ${chalk.grey(config.delimiter)} ${this.name} --id c391b57d-5783-4c53-9236-cefb5c6ef323 --UsageGuidelinesUrl https://contoso.sharepoint.com/sites/compliance --ClassificationList 'HBI, MBI, LBI, GDPR' --DefaultClassification MBI
`);
  }
}

module.exports = new GraphGroupSettingSetCommand();