import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate, CommandError
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../GraphCommand';
import { DirectorySetting, UpdateDirectorySetting } from './DirectorySetting';
import { DirectorySettingValue } from './DirectorySettingValue';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  classifications: string;
  defaultClassification: string;
  usageGuidelinesUrl?: string;
  guestUsageGuidelinesUrl?: string;
}

class GraphSiteClassificationEnableCommand extends GraphCommand {
  public get name(): string {
    return `${commands.SITECLASSIFICATION_ENABLE}`;
  }

  public get description(): string {
    return 'Enables site classification configuration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.usageGuidelinesUrl = typeof args.options.usageGuidelinesUrl !== 'undefined';
    telemetryProps.guestUsageGuidelinesUrl = typeof args.options.guestUsageGuidelinesUrl !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${auth.service.resource}/beta/directorySettingTemplates`,
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
      .then((res: { value: DirectorySetting[]; }): request.RequestPromise | Promise<void> => {
        const unifiedGroupSetting: DirectorySetting[] = res.value.filter((directorySetting: DirectorySetting): boolean => {
          return directorySetting.displayName === 'Group.Unified';
        });

        if (!unifiedGroupSetting ||
          unifiedGroupSetting.length === 0) {
          cb(new CommandError("Missing DirectorySettingTemplate for \"Group.Unified\""));
          return Promise.reject();
        }

        const updatedDirSettings: UpdateDirectorySetting = new UpdateDirectorySetting();
        updatedDirSettings.templateId = unifiedGroupSetting[0].id;

        unifiedGroupSetting[0].values.forEach((directorySetting: DirectorySettingValue) => {
          switch (directorySetting.name) {
            case "ClassificationList":
              updatedDirSettings.values.push({
                "name": directorySetting.name,
                "value": args.options.classifications as string
              });
              break;
            case "DefaultClassification":
              updatedDirSettings.values.push({
                "name": directorySetting.name,
                "value": args.options.defaultClassification as string
              });
              break;
            case "UsageGuidelinesUrl":
              if (args.options.usageGuidelinesUrl) {
                updatedDirSettings.values.push({
                  "name": directorySetting.name,
                  "value": args.options.usageGuidelinesUrl as string
                });
              }
              else {
                updatedDirSettings.values.push({
                  "name": directorySetting.name,
                  "value": directorySetting.defaultValue as string
                })
              }
              break;
            case "GuestUsageGuidelinesUrl":
              if (args.options.guestUsageGuidelinesUrl) {
                updatedDirSettings.values.push({
                  "name": directorySetting.name,
                  "value": args.options.guestUsageGuidelinesUrl as string
                });
              }
              else {
                updatedDirSettings.values.push({
                  "name": directorySetting.name,
                  "value": directorySetting.defaultValue as string
                })
              }
              break;
            default:
              updatedDirSettings.values.push({
                "name": directorySetting.name,
                "value": directorySetting.defaultValue as string
              });
              break;
          }
        });

        const requestOptions: any = {
          url: `${auth.service.resource}/beta/settings`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          }),
          json: true,
          body: updatedDirSettings,
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-c, --classifications <classifications>',
        description: 'Comma-separated list of classifications to enable in the tenant'
      },
      {
        option: '-d, --defaultClassification <defaultClassification>',
        description: 'Classification to use by default'
      },
      {
        option: '-u, --usageGuidelinesUrl [usageGuidelinesUrl]',
        description: 'URL with usage guidelines for members'
      },
      {
        option: '-g, --guestUsageGuidelinesUrl [guestUsageGuidelinesUrl]',
        description: 'URL with usage guidelines for guests'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.classifications) {
        return 'Required option classifications missing';
      }

      if (!args.options.defaultClassification) {
        return 'Required option defaultClassification missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.

    To set the Office 365 Tenant site classification, you have to first login
    to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:
  
    Enable site classification 
      ${chalk.grey(config.delimiter)} ${this.name} --classifications "High, Medium, Low" --defaultClassification "Medium" 

    Enable site classification with a usage guidelines URL 
      ${chalk.grey(config.delimiter)} ${this.name} --classifications "High, Medium, Low" --defaultClassification "Medium" --usageGuidelinesUrl "http://aka.ms/pnp"

    Enable site classification with usage guidelines URLs for guests and members
      ${chalk.grey(config.delimiter)} ${this.name} --classifications "High, Medium, Low" -d "Medium" --usageGuidelinesUrl "http://aka.ms/pnp" --guestUsageGuidelinesUrl "http://aka.ms/pnp" 

  More information:

    SharePoint "modern" sites classification
      https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification
    `);
  }
}

module.exports = new GraphSiteClassificationEnableCommand();