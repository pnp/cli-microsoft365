import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
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

class AadSiteClassificationEnableCommand extends GraphCommand {
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
    const requestOptions: any = {
      url: `${this.resource}/beta/directorySettingTemplates`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    };

    request
      .get<{ value: DirectorySetting[]; }>(requestOptions)
      .then((res: { value: DirectorySetting[]; }): Promise<void> => {
        const unifiedGroupSetting: DirectorySetting[] = res.value.filter((directorySetting: DirectorySetting): boolean => {
          return directorySetting.displayName === 'Group.Unified';
        });

        if (!unifiedGroupSetting ||
          unifiedGroupSetting.length === 0) {
          return Promise.reject("Missing DirectorySettingTemplate for \"Group.Unified\"");
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
          url: `${this.resource}/beta/settings`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          json: true,
          body: updatedDirSettings,
        };

        return request.post(requestOptions);
      })
      .then((): void => {
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
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.

  Examples:
  
    Enable site classification 
      ${this.name} --classifications "High, Medium, Low" --defaultClassification "Medium" 

    Enable site classification with a usage guidelines URL 
      ${this.name} --classifications "High, Medium, Low" --defaultClassification "Medium" --usageGuidelinesUrl "http://aka.ms/pnp"

    Enable site classification with usage guidelines URLs for guests and members
      ${this.name} --classifications "High, Medium, Low" --defaultClassification "Medium" --usageGuidelinesUrl "http://aka.ms/pnp" --guestUsageGuidelinesUrl "http://aka.ms/pnp" 

  More information:

    SharePoint "modern" sites classification
      https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-site-classification
    `);
  }
}

module.exports = new AadSiteClassificationEnableCommand();