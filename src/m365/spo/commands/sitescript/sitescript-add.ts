import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ContextInfo, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  title: z.string().alias('t'),
  content: z.string()
    .refine(val => {
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
    .alias('c'),
  description: z.string().optional().alias('d')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteScriptAddCommand extends SpoCommand {
  public get name(): string {
    return commands.SITESCRIPT_ADD;
  }

  public get description(): string {
    return 'Adds site script for use with site designs';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestDigest: ContextInfo = await spo.getRequestDigest(spoUrl);
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='${formatting.encodeQueryParameter(args.options.title)}'&@description='${formatting.encodeQueryParameter(args.options.description || '')}'`,
        headers: {
          'X-RequestDigest': requestDigest.FormDigestValue,
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        data: JSON.parse(args.options.content),
        responseType: 'json'
      };

      const res: any = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteScriptAddCommand();