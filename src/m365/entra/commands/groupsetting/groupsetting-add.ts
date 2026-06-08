import { GroupSettingTemplate } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  templateId: z.uuid().alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupSettingAddCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTING_ADD;
  }

  public get description(): string {
    return 'Creates a group setting';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving group setting template with id '${args.options.templateId}'...`);
    }

    try {
      let requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groupSettingTemplates/${args.options.templateId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const groupSettingTemplate = await request.get<GroupSettingTemplate>(requestOptions);
      requestOptions = {
        url: `${this.resource}/v1.0/groupSettings`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        data: {
          templateId: args.options.templateId,
          values: this.getGroupSettingValues(args.options, groupSettingTemplate)
        },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getGroupSettingValues(options: any, groupSettingTemplate: GroupSettingTemplate): { name: string; value: string }[] {
    const values: { name: string; value: string }[] = [];
    const excludeOptions: string[] = [
      'templateId',
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

    groupSettingTemplate.values!.forEach(v => {
      if (!values.find(e => e.name === v.name)) {
        values.push({
          name: v.name!,
          value: v.defaultValue!
        });
      }
    });

    return values;
  }
}

export default new EntraGroupSettingAddCommand();