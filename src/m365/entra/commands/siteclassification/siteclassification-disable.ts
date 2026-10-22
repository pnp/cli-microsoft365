import { GroupSetting } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  force: z.boolean().optional().alias('f')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraSiteClassificationDisableCommand extends GraphCommand {
  public get name(): string {
    return commands.SITECLASSIFICATION_DISABLE;
  }

  public get description(): string {
    return 'Disables site classification';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const disableSiteClassification = async (): Promise<void> => {
      try {
        let requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groupSettings`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        const res = await request.get<{ value: GroupSetting[] }>(requestOptions);
        if (res.value.length === 0) {
          throw 'Site classification is not enabled.';
        }

        const unifiedGroupSetting: GroupSetting[] = res.value.filter((directorySetting: GroupSetting): boolean => {
          return directorySetting.displayName === 'Group.Unified';
        });

        if (!unifiedGroupSetting || unifiedGroupSetting.length === 0) {
          throw 'Missing DirectorySettingTemplate for "Group.Unified"';
        }

        if (!unifiedGroupSetting[0] ||
          !unifiedGroupSetting[0].id || unifiedGroupSetting[0].id.length === 0) {
          throw 'Missing UnifiedGroupSettting id';
        }

        requestOptions = {
          url: `${this.resource}/v1.0/groupSettings/` + unifiedGroupSetting[0].id,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await disableSiteClassification();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to disable site classification?` });

      if (result) {
        await disableSiteClassification();
      }
    }
  }
}

export default new EntraSiteClassificationDisableCommand();