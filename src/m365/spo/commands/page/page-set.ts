import { z } from 'zod';
import { Auth } from '../../../../Auth.js';
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
  layoutType: z.enum(supportedPageLayouts).optional().alias('l'),
  promoteAs: z.enum(supportedPromoteAs).optional().alias('p'),
  demoteFrom: z.enum(['NewsPage']).optional(),
  commentsEnabled: z.boolean().optional(),
  publish: z.boolean().optional(),
  publishMessage: z.string().optional(),
  description: z.string().optional(),
  title: z.string().optional(),
  content: z.string()
    .refine(content => {
      try {
        JSON.parse(content);
        return true;
      }
      catch {
        return false;
      }
    }, {
      error: 'Specified content is not a valid JSON string.'
    })
    .optional(),
  isRetired: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageSetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_SET;
  }

  public get description(): string {
    return 'Updates modern page properties';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !!(opts.layoutType || opts.promoteAs || opts.demoteFrom || opts.commentsEnabled !== undefined ||
        opts.publish || opts.description || opts.title || opts.content || opts.isRetired !== undefined), {
        error: 'Specify at least one option to update.'
      })
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
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let pageName: string = args.options.name;
    const fileNameWithoutExtension: string = pageName.replace('.aspx', '');

    let pageTitle: string = args.options.title || "";
    let pageDescription: string = args.options.description || "";
    const pageData: any = {};

    if (!pageName.endsWith('.aspx')) {
      pageName += '.aspx';
    }
    const listServerRelativeUrl = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages`;
    const serverRelativeFileUrl: string = `${listServerRelativeUrl}/${pageName}`;

    const listUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, listServerRelativeUrl);
    const requestUrl = `${args.options.webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;

    const needsToSavePage = !!args.options.title || !!args.options.description || !!args.options.content;

    try {
      const requestDigestResult = await spo.getRequestDigest(args.options.webUrl);
      const requestDigest = requestDigestResult.FormDigestValue;
      const page = await Page.checkout(args.options.name, args.options.webUrl, logger, this.verbose);

      pageTitle = pageTitle || page.Title;
      let pageId: number | null = page.Id;

      const bannerImageUrl = page.BannerImageUrl;
      const canvasContent1 = args.options.content || page.CanvasContent1;
      const layoutWebpartsContent = page.LayoutWebpartsContent;
      pageDescription = pageDescription || page.Description;
      const topicHeader = page.TopicHeader;
      const authorByline = page.AuthorByline;

      if (args.options.layoutType) {
        const listItemSetOptions: any = {
          PageLayoutType: args.options.layoutType
        };
        if (args.options.layoutType === 'Article') {
          listItemSetOptions.PromotedState = 0;
          listItemSetOptions.BannerImageUrl = `${resource}/_layouts/15/images/sitepagethumbnail.png, /_layouts/15/images/sitepagethumbnail.png`;
        }
        await spo.systemUpdateListItem(requestUrl, pageId, logger, this.verbose, listItemSetOptions);
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
              PromotedState: 2,
              FirstPublishedDate: new Date().toISOString()
            };
            await spo.systemUpdateListItem(requestUrl, pageId, logger, this.verbose, listItemSetOptions);
            break;
          }
          case 'Template': {
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              'X-HTTP-Method': 'POST',
              'IF-MATCH': '*',
              accept: 'application/json;odata=nometadata'
            };
            requestOptions.url = `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsTemplate`;
            const res = await request.post<{ Id: number | null, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string }>(requestOptions);
            if (fileNameWithoutExtension) {
              pageData.Title = fileNameWithoutExtension;
            }
            if (pageDescription) {
              pageData.Description = pageDescription;
            }
            if (res.BannerImageUrl) {
              pageData.BannerImageUrl = res.BannerImageUrl;
            }
            if (res.LayoutWebpartsContent) {
              pageData.LayoutWebpartsContent = res.LayoutWebpartsContent;
            }
            if (res.CanvasContent1) {
              pageData.CanvasContent1 = res.CanvasContent1;
            }
            pageId = res.Id;
            break;
          }
        }
      }

      if (args.options.promoteAs !== 'Template') {
        if (pageTitle) {
          pageData.Title = pageTitle;
        }
        if (pageDescription) {
          pageData.Description = pageDescription;
        }
        if (bannerImageUrl) {
          pageData.BannerImageUrl = bannerImageUrl;
        }
        if (canvasContent1) {
          pageData.CanvasContent1 = canvasContent1;
        }
        if (layoutWebpartsContent) {
          pageData.LayoutWebpartsContent = layoutWebpartsContent;
        }
        if (topicHeader) {
          pageData.TopicHeader = topicHeader;
        }
        if (authorByline) {
          pageData.AuthorByline = authorByline;
        }
      }

      if (needsToSavePage) {
        const requestOptions: CliRequestOptions = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePage`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: pageData
        };

        await request.post(requestOptions);
      }

      if (args.options.promoteAs === 'Template') {
        const requestOptions: CliRequestOptions = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: pageData
        };

        await request.post(requestOptions);
      }

      if (typeof args.options.commentsEnabled !== 'undefined') {
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${serverRelativeFileUrl}')/ListItemAllFields/SetCommentsDisabled(${args.options.commentsEnabled === false})`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }

      if (args.options.demoteFrom === 'NewsPage') {
        const listItemSetOptions: any = {
          PromotedState: 0
        };
        await spo.systemUpdateListItem(requestUrl, pageId!, logger, this.verbose, listItemSetOptions);
      }

      if (args.options.isRetired !== undefined) {
        await spo.systemUpdateListItem(requestUrl, pageId!, logger, this.verbose, {
          _SPIsRetired: args.options.isRetired
        });
      }

      let requestOptions: CliRequestOptions;

      if (!args.options.publish) {
        if (args.options.promoteAs === 'Template' || !pageId) {
          return;
        }

        requestOptions = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            'accept': 'application/json;odata=nometadata'
          },
          data: pageData
        };
      }
      else {
        requestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${serverRelativeFileUrl}')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${formatting.encodeQueryParameter(args.options.publishMessage || '')}'&@a2=1`,
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

export default new SpoPageSetCommand();