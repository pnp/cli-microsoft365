import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { ContextInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(val => validation.isValidGuid(val), {
    message: 'The value must be a valid GUID.'
  }).alias('i'),
  title: z.string().optional().alias('t'),
  description: z.string().optional().alias('d'),
  version: z.string().optional()
    .refine(val => val === undefined || !isNaN(parseInt(val)), {
      message: 'Version must be a number.'
    })
    .alias('v'),
  content: z.string().optional()
    .refine(val => {
      if (val === undefined) {
        return true;
      }
      try {
        JSON.parse(val);
        return true;
      }
      catch {
        return false;
      }
    }, {
      message: 'Specified content value is not a valid JSON string.'
    })
    .alias('c')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteScriptSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITESCRIPT_SET;
  }

  public get description(): string {
    return 'Updates existing site script';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const formDigest: ContextInfo = await spo.getRequestDigest(spoUrl);
      const updateInfo: any = {
        Id: args.options.id
      };
      if (args.options.title) {
        updateInfo.Title = args.options.title;
      }
      if (args.options.description) {
        updateInfo.Description = args.options.description;
      }
      if (args.options.version) {
        updateInfo.Version = parseInt(args.options.version);
      }
      if (args.options.content) {
        updateInfo.Content = args.options.content;
      }

      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`,
        headers: {
          'X-RequestDigest': formDigest.FormDigestValue,
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        data: { updateInfo: updateInfo },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteScriptSetCommand();