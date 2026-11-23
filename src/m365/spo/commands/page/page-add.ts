import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page, supportedPageLayouts, supportedPromoteAs } from './Page.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  name: z.string().alias('n'),
  title: z.string().optional().alias('t'),
  layoutType: z.enum(supportedPageLayouts).optional().alias('l'),
  promoteAs: z.enum(supportedPromoteAs).optional().alias('p'),
  commentsEnabled: z.boolean().optional(),
  publish: z.boolean().optional(),
  publishMessage: z.string().optional(),
  description: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageAddCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_ADD;
  }

  public get description(): string {
    return 'Creates modern page';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.promoteAs !== 'HomePage' || opts.layoutType === 'Home', {
        error: 'You can only promote home pages as site home page',
        path: ['promoteAs']
      })
      .refine(opts => opts.promoteAs !== 'NewsPage' || !opts.layoutType || opts.layoutType === 'Article', {
        error: 'You can only promote article pages as news article',
        path: ['promoteAs']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let pageName: string = args.options.name;
    const fileNameWithoutExtension: string = pageName.replace('.aspx', '');
    let bannerImageUrl: string = '';
    let canvasContent1: string = '';
    let layoutWebpartsContent: string = '';
    const pageTitle: string = args.options.title ? args.options.title : (args.options.name.indexOf('.aspx') > -1 ? args.options.name.substring(0, args.options.name.indexOf('.aspx')) : args.options.name);
    let pageId: number | null = null;
    const pageDescription: string = args.options.description || "";

    if (!pageName.endsWith('.aspx')) {
      pageName += '.aspx';
    }
    const listServerRelativeUrl = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages`;
    const serverRelativeFileUrl: string = `${listServerRelativeUrl}/${pageName}`;

    try {
      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      const requestDigest = reqDigest.FormDigestValue;

      let requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/sitepages/pages`,
        responseType: 'json',
        headers: {
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        data: {
          PageLayoutType: args.options.layoutType || 'Article',
          Name: pageName,
          PromotedState: args.options.promoteAs === 'NewsPage' ? 2 : 0,
          Title: pageTitle
        }
      };

      const template = await request.post<{ UniqueId: string, Id: string }>(requestOptions);
      const itemId = template.UniqueId;
      const file = await spo.getFileAsListItemByUrl(args.options.webUrl, serverRelativeFileUrl, logger, this.verbose);
      const listItemId = file.Id;

      const pageProps = await Page.checkout(pageName, args.options.webUrl, logger, this.verbose);
      if (pageProps) {
        pageId = pageProps.Id;

        bannerImageUrl = pageProps.BannerImageUrl;
        canvasContent1 = pageProps.CanvasContent1;
        layoutWebpartsContent = pageProps.LayoutWebpartsContent;
      }

      if (args.options.promoteAs) {
        const requestOptions: CliRequestOptions = {
          responseType: 'json'
        };

        switch (args.options.promoteAs) {
          case 'HomePage':
            requestOptions.url = `${args.options.webUrl}/_api/web/rootfolder`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };
            requestOptions.data = {
              WelcomePage: `SitePages/${pageName}`
            };
            await request.post(requestOptions);
            break;
          case 'NewsPage': {
            const listItemSetOptions: any = {
              FirstPublishedDate: new Date().toISOString()
            };
            const listUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, listServerRelativeUrl);
            const requestUrl = `${args.options.webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;

            await spo.systemUpdateListItem(requestUrl, listItemId, logger, this.verbose, listItemSetOptions);
            break;
          }
          case 'Template': {
            requestOptions.url = `${args.options.webUrl}/_api/SitePages/Pages(${listItemId})/SavePageAsTemplate`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              'X-HTTP-Method': 'POST',
              'IF-MATCH': '*',
              accept: 'application/json;odata=nometadata'
            };

            const tmpl = await request.post<{ Id: number | null, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string, UniqueId: string }>(requestOptions);

            bannerImageUrl = tmpl.BannerImageUrl;
            canvasContent1 = tmpl.CanvasContent1;
            layoutWebpartsContent = tmpl.LayoutWebpartsContent;
            pageId = tmpl.Id;

            requestOptions.url = `${args.options.webUrl}/_api/web/getfilebyid('${tmpl.UniqueId}')/ListItemAllFields/SetCommentsDisabled(${!args.options.commentsEnabled})`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };

            await request.post(requestOptions);
            break;
          }
        }
      }

      requestOptions = {
        responseType: 'json',
        url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePage`,
        headers: {
          'X-RequestDigest': requestDigest,
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        data: {
          BannerImageUrl: bannerImageUrl,
          CanvasContent1: canvasContent1,
          LayoutWebpartsContent: layoutWebpartsContent,
          Description: pageDescription
        }
      };

      await request.post(requestOptions);

      if (args.options.promoteAs === 'Template') {
        const requestOptions: any = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            Title: fileNameWithoutExtension,
            BannerImageUrl: bannerImageUrl,
            CanvasContent1: canvasContent1,
            LayoutWebpartsContent: layoutWebpartsContent,
            Description: pageDescription
          }
        };

        await request.post(requestOptions);
      }

      requestOptions = {
        url: `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/ListItemAllFields/SetCommentsDisabled(${!args.options.commentsEnabled})`,
        headers: {
          'X-RequestDigest': requestDigest,
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);

      if (!args.options.publish) {
        if (args.options.promoteAs !== 'Template' && pageId) {
          requestOptions = {
            responseType: 'json',
            url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
            headers: {
              'content-type': 'application/json;odata=nometadata',
              'accept': 'application/json;odata=nometadata'
            },
            data: {
              Title: pageTitle,
              Description: pageDescription,
              BannerImageUrl: bannerImageUrl,
              CanvasContent1: canvasContent1,
              LayoutWebpartsContent: layoutWebpartsContent
            }
          };
        }
      }
      else {
        requestOptions = {
          url: `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${formatting.encodeQueryParameter(args.options.publishMessage || '').replace(/'/g, '%39')}'&@a2=1`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
      }

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageAddCommand();
