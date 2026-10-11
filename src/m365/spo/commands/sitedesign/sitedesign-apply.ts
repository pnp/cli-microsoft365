import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), { message: 'The value must be a valid GUID.' }).alias('i'),
  webUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, { message: 'The value is not a valid SharePoint Online site URL.' }).alias('u'),
  asTask: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteDesignApplyCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_APPLY;
  }

  public get description(): string {
    return 'Applies a site design to an existing site collection';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestBody: any = {
        siteDesignId: args.options.id,
        webUrl: args.options.webUrl
      };

      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.${args.options.asTask ? 'AddSiteDesignTask' : 'ApplySiteDesign'}`,
        headers: {
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const res: any = await request.post(requestOptions);

      if (res.value) {
        await logger.log(res.value);
      }
      else {
        await logger.log(res);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteDesignApplyCommand();