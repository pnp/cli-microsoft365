import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import { DirectorySetting, UpdateDirectorySetting } from './DirectorySetting';
import { DirectorySettingValue } from './DirectorySettingValue';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  classifications?: string;
  defaultClassification?: string;
  usageGuidelinesUrl?: string;
  guestUsageGuidelinesUrl?: string;
}

class AadSiteClassificationUpdateCommand extends GraphCommand {
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
    const requestOptions: any = {
      url: `${this.resource}/beta/settings`,
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
          url: `${this.resource}/beta/settings/${unifiedGroupSetting[0].id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          json: true,
          body: updatedDirSettings,
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
}

module.exports = new AadSiteClassificationUpdateCommand();