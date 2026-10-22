import { GroupSetting, SettingValue } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  classifications: z.string().alias('c'),
  defaultClassification: z.string().alias('d'),
  usageGuidelinesUrl: z.string().optional().alias('u'),
  guestUsageGuidelinesUrl: z.string().optional().alias('g')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraSiteClassificationEnableCommand extends GraphCommand {
  public get name(): string {
    return commands.SITECLASSIFICATION_ENABLE;
  }

  public get description(): string {
    return 'Enables site classification configuration';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groupSettingTemplates`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: GroupSetting[]; }>(requestOptions);

      const unifiedGroupSetting: GroupSetting[] = res.value.filter((directorySetting: GroupSetting): boolean => {
        return directorySetting.displayName === 'Group.Unified';
      });

      if (!unifiedGroupSetting ||
        unifiedGroupSetting.length === 0) {
        throw "Missing DirectorySettingTemplate for \"Group.Unified\"";
      }

      const updatedDirSettings: GroupSetting = { values: [], templateId: unifiedGroupSetting[0].id } as GroupSetting;

      unifiedGroupSetting[0].values!.forEach((directorySetting: SettingValue) => {
        switch (directorySetting.name) {
          case "ClassificationList":
            updatedDirSettings.values!.push({
              "name": directorySetting.name,
              "value": args.options.classifications as string
            });
            break;
          case "DefaultClassification":
            updatedDirSettings.values!.push({
              "name": directorySetting.name,
              "value": args.options.defaultClassification as string
            });
            break;
          case "UsageGuidelinesUrl":
            if (args.options.usageGuidelinesUrl) {
              updatedDirSettings.values!.push({
                "name": directorySetting.name,
                "value": args.options.usageGuidelinesUrl as string
              });
            }
            else {
              updatedDirSettings.values!.push({
                "name": directorySetting.name,
                "value": (directorySetting as any).defaultValue as string
              });
            }
            break;
          case "GuestUsageGuidelinesUrl":
            if (args.options.guestUsageGuidelinesUrl) {
              updatedDirSettings.values!.push({
                "name": directorySetting.name,
                "value": args.options.guestUsageGuidelinesUrl as string
              });
            }
            else {
              updatedDirSettings.values!.push({
                "name": directorySetting.name,
                "value": (directorySetting as any).defaultValue as string
              });
            }
            break;
          default:
            updatedDirSettings.values!.push({
              "name": directorySetting.name,
              "value": (directorySetting as any).defaultValue as string
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

export default new EntraSiteClassificationEnableCommand();