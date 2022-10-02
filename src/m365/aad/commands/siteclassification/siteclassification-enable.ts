import { Logger } from '../../../../cli/Logger';
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
  classifications: string;
  defaultClassification: string;
  usageGuidelinesUrl?: string;
  guestUsageGuidelinesUrl?: string;
}

class AadSiteClassificationEnableCommand extends GraphCommand {
  public get name(): string {
    return commands.SITECLASSIFICATION_ENABLE;
  }

  public get description(): string {
    return 'Enables site classification configuration';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        usageGuidelinesUrl: typeof args.options.usageGuidelinesUrl !== 'undefined',
        guestUsageGuidelinesUrl: typeof args.options.guestUsageGuidelinesUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-c, --classifications <classifications>'
      },
      {
        option: '-d, --defaultClassification <defaultClassification>'
      },
      {
        option: '-u, --usageGuidelinesUrl [usageGuidelinesUrl]'
      },
      {
        option: '-g, --guestUsageGuidelinesUrl [guestUsageGuidelinesUrl]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let requestOptions: any = {
        url: `${this.resource}/v1.0/groupSettingTemplates`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: DirectorySetting[]; }>(requestOptions);

      const unifiedGroupSetting: DirectorySetting[] = res.value.filter((directorySetting: DirectorySetting): boolean => {
        return directorySetting.displayName === 'Group.Unified';
      });

      if (!unifiedGroupSetting ||
        unifiedGroupSetting.length === 0) {
        throw "Missing DirectorySettingTemplate for \"Group.Unified\"";
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
                "value": directorySetting.defaultValue as string
              });
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

      requestOptions = {
        url: `${this.resource}/v1.0/groupSettings`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: updatedDirSettings
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadSiteClassificationEnableCommand();