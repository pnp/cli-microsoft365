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
import { DirectorySetting, UpdateDirectorySetting } from './DirectorySetting';
import { DirectorySettingValue } from './DirectorySettingValue';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  classifications?: string;
  defaultClassification?: string;
  usageGuidelinesUrl?: string;
  guestUsageGuidelinesUrl?: string;
}

class GraphSiteClassificationUpdateCommand extends GraphCommand {
  public get name(): string {
    return `${commands.SITECLASSIFICATION_SET}`;
  }

  public get description(): string {
    return 'Updates site classification configuration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.classifications = typeof args.options.classifications !== 'undefined';
    telemetryProps.defaultClassification = typeof args.options.defaultClassification !== 'undefined';
    telemetryProps.usageGuidelinesUrl = typeof args.options.usageGuidelinesUrl !== 'undefined';
    telemetryProps.guestUsageGuidelinesUrl = typeof args.options.guestUsageGuidelinesUrl !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${auth.service.resource}/beta/settings`,
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
          return Promise.reject("There is no previous defined site classification which can updated.");
        }

        const updatedDirSettings: UpdateDirectorySetting = new UpdateDirectorySetting();

        unifiedGroupSetting[0].values.forEach((directorySetting: DirectorySettingValue) => {
          switch (directorySetting.name) {
            case "ClassificationList":
              if (args.options.classifications) {
                updatedDirSettings.values.push({
                  "name": directorySetting.name,
                  "value": args.options.classifications as string
                });
              }
              else {
                updatedDirSettings.values.push({
                  "name": directorySetting.name,
                  "value": directorySetting.value as string
                });
              }
              break;
            case "DefaultClassification":
              if (args.options.defaultClassification) {
                updatedDirSettings.values.push({
                  "name": directorySetting.name,
                  "value": args.options.defaultClassification as string
                });
              }
              else {
                updatedDirSettings.values.push({
                  "name": directorySetting.name,
                  "value": directorySetting.value as string
                });
              }
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
                  "value": directorySetting.value as string
                });
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
                  "value": directorySetting.value as string
                });
              }
              break;
            default:
              updatedDirSettings.values.push({
                "name": directorySetting.name,
                "value": directorySetting.value as string
              });
              break;
          }
        });

        const requestOptions: any = {
          url: `${auth.service.resource}/beta/settings/${unifiedGroupSetting[0].id}`,
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

        if (this.verbose) {
          cmd.log('The updated classification settings will be:');
          cmd.log(updatedDirSettings);
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-c, --classifications [classifications]',
        description: 'Comma-separated list of classifications'
      },
      {
        option: '-d, --defaultClassification [defaultClassification]',
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
      if (!args.options.classifications &&
        !args.options.defaultClassification &&
        !args.options.usageGuidelinesUrl &&
        !args.options.guestUsageGuidelinesUrl) {
        return 'Specify at least one property to update';
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

    To update the Office 365 Tenant site classification configuration, you have
    to first login to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    Update Office 365 Tenant site classification configuration
      ${chalk.grey(config.delimiter)} ${this.name} --classifications "High, Medium, Low" --defaultClassification "Medium" 

    Update only the default classification
      ${chalk.grey(config.delimiter)} ${this.name} --defaultClassification "Low"

    Update site classification with a usage guidelines URL 
      ${chalk.grey(config.delimiter)} ${this.name} --usageGuidelinesUrl "http://aka.ms/pnp"

    Update site classification with usage guidelines URLs for guests and members
      ${chalk.grey(config.delimiter)} ${this.name} --usageGuidelinesUrl "http://aka.ms/pnp" --guestUsageGuidelinesUrl "http://aka.ms/pnp" 

  More information:

    SharePoint "modern" sites classification
      https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification
    `);
  }
}

module.exports = new GraphSiteClassificationUpdateCommand();