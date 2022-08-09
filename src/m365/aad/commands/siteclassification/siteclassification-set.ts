import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { DirectorySetting, UpdateDirectorySetting } from './DirectorySetting';
import { DirectorySettingValue } from './DirectorySettingValue';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  classifications?: string;
  defaultClassification?: string;
  usageGuidelinesUrl?: string;
  guestUsageGuidelinesUrl?: string;
}

class AadSiteClassificationSetCommand extends GraphCommand {
  public get name(): string {
    return commands.SITECLASSIFICATION_SET;
  }

  public get description(): string {
    return 'Updates site classification configuration';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        classifications: typeof args.options.classifications !== 'undefined',
        defaultClassification: typeof args.options.defaultClassification !== 'undefined',
        usageGuidelinesUrl: typeof args.options.usageGuidelinesUrl !== 'undefined',
        guestUsageGuidelinesUrl: typeof args.options.guestUsageGuidelinesUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-c, --classifications [classifications]'
      },
      {
        option: '-d, --defaultClassification [defaultClassification]'
      },
      {
        option: '-u, --usageGuidelinesUrl [usageGuidelinesUrl]'
      },
      {
        option: '-g, --guestUsageGuidelinesUrl [guestUsageGuidelinesUrl]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.classifications &&
          !args.options.defaultClassification &&
          !args.options.usageGuidelinesUrl &&
          !args.options.guestUsageGuidelinesUrl) {
          return 'Specify at least one property to update';
        }
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/groupSettings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
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
          url: `${this.resource}/v1.0/groupSettings/${unifiedGroupSetting[0].id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: updatedDirSettings
        };

        return request.patch(requestOptions);
      })
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadSiteClassificationSetCommand();