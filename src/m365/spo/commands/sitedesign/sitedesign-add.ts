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
  title: z.string().alias('t'),
  webTemplate: z.string().pipe(z.enum(['TeamSite', 'CommunicationSite'])).alias('w'),
  siteScripts: z.string().superRefine((val, ctx) => {
    const items = val.split(',').map(s => s.trim());
    const invalid = items.filter(item => !validation.isValidGuid(item));
    if (invalid.length > 0) {
      ctx.addIssue({
        code: 'custom',
        message: `The following values are not valid GUIDs: ${invalid.join(', ')}`
      });
    }
  }).alias('s'),
  description: z.string().optional().alias('d'),
  previewImageUrl: z.string().optional().alias('m'),
  previewImageAltText: z.string().optional().alias('a'),
  thumbnailUrl: z.string().optional(),
  isDefault: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteDesignAddCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_ADD;
  }

  public get description(): string {
    return 'Adds site design for creating modern sites';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestDigest: ContextInfo = await spo.getRequestDigest(spoUrl);
      const info: any = {
        Title: args.options.title,
        WebTemplate: args.options.webTemplate === 'TeamSite' ? '64' : '68',
        SiteScriptIds: args.options.siteScripts.split(',').map(i => i.trim())
      };

      if (args.options.description) {
        info.Description = args.options.description;
      }
      if (args.options.previewImageUrl) {
        info.PreviewImageUrl = args.options.previewImageUrl;
      }
      if (args.options.previewImageAltText) {
        info.PreviewImageAltText = args.options.previewImageAltText;
      }
      if (args.options.thumbnailUrl) {
        info.ThumbnailUrl = args.options.thumbnailUrl;
      }
      if (args.options.isDefault) {
        info.IsDefault = true;
      }

      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`,
        headers: {
          'X-RequestDigest': requestDigest.FormDigestValue,
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        data: { info: info },
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

export default new SpoSiteDesignAddCommand();
