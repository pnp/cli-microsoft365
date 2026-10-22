import { GroupSettingTemplate } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  displayName: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupSettingTemplateGetCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTINGTEMPLATE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Entra group settings template';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.displayName].filter(Boolean).length === 1, {
        error: 'Specify either id or displayName',
        params: {
          customCode: 'optionSet',
          options: ['id', 'displayName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const templates = await odata.getAllItems<GroupSettingTemplate>(`${this.resource}/v1.0/groupSettingTemplates`);

      const groupSettingTemplate: GroupSettingTemplate[] = templates.filter(t => args.options.id ? t.id === args.options.id : t.displayName === args.options.displayName);

      if (groupSettingTemplate && groupSettingTemplate.length > 0) {
        await logger.log(groupSettingTemplate.pop());
      }
      else {
        throw `Resource '${(args.options.id || args.options.displayName)}' does not exist.`;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraGroupSettingTemplateGetCommand();