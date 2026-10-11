import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, { message: 'The value is not a valid SharePoint Online site URL.' }).alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteDesignTaskListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_TASK_LIST;
  }

  public get description(): string {
    return 'Lists site designs scheduled for execution on the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['ID', 'SiteDesignID', 'LogonName'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTasks`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res: { value: any[] } = await request.post<{ value: any[] }>(requestOptions);
      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteDesignTaskListCommand();