import { GroupSetting, SettingValue } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  classifications: z.string().optional().alias('c'),
  defaultClassification: z.string().optional().alias('d'),
  usageGuidelinesUrl: z.string().optional().alias('u'),
  guestUsageGuidelinesUrl: z.string().optional().alias('g')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraSiteClassificationSetCommand extends GraphCommand {
  public get name(): string {
    return commands.SITECLASSIFICATION_SET;
  }

  public get description(): string {
    return 'Updates site classification configuration';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => options.classifications || options.defaultClassification || options.usageGuidelinesUrl || options.guestUsageGuidelinesUrl, {
        error: 'Specify at least one property to update',
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groupSettings`,
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
        throw "There is no previous defined site classification which can updated.";
      }

      const updatedDirSettings: GroupSetting = { values: [] } as GroupSetting;

      unifiedGroupSetting[0]!.values!.forEach((directorySetting: SettingValue) => {
        switch (directorySetting.name) {
          case "ClassificationList":
            if (args.options.classifications) {
              updatedDirSettings!.values!.push({
                name: directorySetting.name,
                value: args.options.classifications as string
              });
            }
            else {
              updatedDirSettings!.values!.push({
                name: directorySetting.name,
                value: directorySetting.value as string
              });
            }
            break;
          case "DefaultClassification":
            if (args.options.defaultClassification) {
              updatedDirSettings!.values!.push({
                name: directorySetting.name,
                value: args.options.defaultClassification as string
              });
            }
            else {
              updatedDirSettings!.values!.push({
                name: directorySetting.name,
                value: directorySetting.value as string
              });
            }
            break;
          case "UsageGuidelinesUrl":
            if (args.options.usageGuidelinesUrl) {
              updatedDirSettings!.values!.push({
                name: directorySetting.name,
                value: args.options.usageGuidelinesUrl as string
              });
            }
            else {
              updatedDirSettings!.values!.push({
                name: directorySetting.name,
                value: directorySetting.value as string
              });
            }
            break;
          case "GuestUsageGuidelinesUrl":
            if (args.options.guestUsageGuidelinesUrl) {
              updatedDirSettings!.values!.push({
                name: directorySetting.name,
                value: args.options.guestUsageGuidelinesUrl as string
              });
            }
            else {
              updatedDirSettings!.values!.push({
                name: directorySetting.name,
                value: directorySetting.value as string
              });
            }
            break;
          default:
            updatedDirSettings!.values!.push({
              name: directorySetting.name,
              value: directorySetting.value as string
            });
            break;
        }
      });

      requestOptions = {
        url: `${this.resource}/v1.0/groupSettings/${unifiedGroupSetting[0].id}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: updatedDirSettings
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraSiteClassificationSetCommand();