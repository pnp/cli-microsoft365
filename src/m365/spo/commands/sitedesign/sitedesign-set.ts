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
  title: z.string().optional().alias('t'),
  webTemplate: z.enum(['TeamSite', 'CommunicationSite']).optional().alias('w'),
  siteScripts: z.string().superRefine((val, ctx) => {
    const items = val.split(',').map(s => s.trim());
    const invalid = items.filter(item => !validation.isValidGuid(item));
    if (invalid.length > 0) {
      ctx.addIssue({
        code: 'custom',
        message: `The following values are not valid GUIDs: ${invalid.join(', ')}`
      });
    }
  }).optional().alias('s'),
  description: z.string().optional().alias('d'),
  previewImageUrl: z.string().optional().alias('m'),
  previewImageAltText: z.string().optional().alias('a'),
  thumbnailUrl: z.string().optional(),
  version: z.string().refine(val => !isNaN(Number(val)), { message: 'Version must be a number' }).optional().alias('v'),
  isDefault: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteDesignSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_SET;
  }

  public get description(): string {
    return 'Updates a site design with new values';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const updateInfo: any = {
        Id: args.options.id
      };

      if (args.options.title) {
        updateInfo.Title = args.options.title;
      }
      if (args.options.description) {
        updateInfo.Description = args.options.description;
      }
      if (args.options.siteScripts) {
        updateInfo.SiteScriptIds = args.options.siteScripts.split(',').map(i => i.trim());
      }
      if (args.options.previewImageUrl) {
        updateInfo.PreviewImageUrl = args.options.previewImageUrl;
      }
      if (args.options.previewImageAltText) {
        updateInfo.PreviewImageAltText = args.options.previewImageAltText;
      }
      if (args.options.thumbnailUrl) {
        updateInfo.ThumbnailUrl = args.options.thumbnailUrl;
      }
      if (args.options.webTemplate) {
        updateInfo.WebTemplate = args.options.webTemplate === 'TeamSite' ? '64' : '68';
      }
      if (args.options.version) {
        updateInfo.Version = Number(args.options.version);
      }
      if (typeof args.options.isDefault !== 'undefined') {
        updateInfo.IsDefault = args.options.isDefault;
      }

      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`,
        headers: {
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata',
          responseType: 'json'
        },
        data: { updateInfo: updateInfo },
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

export default new SpoSiteDesignSetCommand();
