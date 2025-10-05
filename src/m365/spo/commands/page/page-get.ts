import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  name: z.string().optional().alias('n'),
  default: z.boolean().optional(),
  metadataOnly: z.boolean().optional()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageGetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_GET;
  }

  public get description(): string {
    return 'Gets information about the specific modern page';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.name, options.default].filter(x => x !== undefined).length === 1, {
        error: `Specify either name or default, but not both.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about the page...`);
    }

    let pageName: string = '';
    try {
      if (args.options.name) {
        pageName = args.options.name.endsWith('.aspx')
          ? args.options.name
          : `${args.options.name}.aspx`;
      }
      else if (args.options.default) {
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/Web/RootFolder?$select=WelcomePage`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const { WelcomePage } = await request.get<{ WelcomePage: string }>(requestOptions);
        pageName = WelcomePage.split('/').pop()!;
      }

      let requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${formatting.encodeQueryParameter(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId,ListItemAllFields/PageLayoutType,ListItemAllFields/CommentsDisabled`,
        headers: {
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const page = await request.get<any>(requestOptions);

      if (page.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
        throw `Page ${pageName} is not a modern page.`;
      }

      let pageItemData: any = {};
      pageItemData = Object.assign({}, page);
      pageItemData.commentsDisabled = page.ListItemAllFields.CommentsDisabled;
      pageItemData.title = page.ListItemAllFields.Title;

      if (page.ListItemAllFields.PageLayoutType) {
        pageItemData.layoutType = page.ListItemAllFields.PageLayoutType;
      }

      if (!args.options.metadataOnly) {
        requestOptions = {
          url: `${args.options.webUrl}/_api/SitePages/Pages(${page.ListItemAllFields.Id})`,
          headers: {
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const res = await request.get<{ CanvasContent1: string }>(requestOptions);
        const canvasData: any[] = JSON.parse(res.CanvasContent1);
        pageItemData.canvasContentJson = res.CanvasContent1;
        if (canvasData && canvasData.length > 0) {
          pageItemData.numControls = canvasData.length;
          const sections = [...new Set(canvasData.filter(c => c.position).map(c => c.position.zoneIndex))];
          pageItemData.numSections = sections.length;
        }
      }

      delete pageItemData.ListItemAllFields.ID;

      await logger.log(pageItemData);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageGetCommand();
