import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SiteDesignTask } from './SiteDesignTask.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), { message: 'The value must be a valid GUID.' }).alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteDesignTaskGetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_TASK_GET;
  }

  public get description(): string {
    return 'Gets information about the specified site design scheduled for execution';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTask`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: {
          taskId: args.options.id
        },
        responseType: 'json'
      };

      const res: SiteDesignTask = await request.post(requestOptions);
      if (!res["odata.null"]) {
        await logger.log(res);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteDesignTaskGetCommand();