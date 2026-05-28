import { GroupSetting } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = z.looseObject({
  ...globalOptionsZod.shape,
  id: z.uuid().alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupSettingSetCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTING_SET;
  }

  public get description(): string {
    return 'Updates the particular group setting';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving group setting with id '${args.options.id}'...`);
    }

    try {
      let requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const groupSetting = await request.get<GroupSetting>(requestOptions);

      requestOptions = {
        url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        data: {
          displayName: groupSetting.displayName,
          templateId: groupSetting.templateId,
          values: this.getGroupSettingValues(args.options, groupSetting)
        },
        responseType: 'json'
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getGroupSettingValues(options: any, groupSetting: GroupSetting): { name: string; value: string }[] {
    const values: { name: string; value: string }[] = [];
    const excludeOptions: string[] = [
      'id',
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

    groupSetting.values!.forEach(v => {
      if (!values.find(e => e.name === v.name)) {
        values.push({
          name: v.name!,
          value: v.value!
        });
      }
    });

    return values;
  }
}

export default new EntraGroupSettingSetCommand();