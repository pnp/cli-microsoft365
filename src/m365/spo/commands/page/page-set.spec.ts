import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command, { options } from './page-set.js';
import { ListItemInstance } from '../listitem/ListItemInstance.js';

describe(commands.PAGE_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const fileResponse = {
    FileSystemObjectType: 0,
    Id: 6,
    ServerRedirectedEmbedUri: 'https://contoso.sharepoint.com/sites/sales/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview',
    ServerRedirectedEmbedUrl: 'https://contoso.sharepoint.com/sites/sales/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview',
    ContentTypeId: '0x0101008E462E3ACE8DB844B3BEBF9473311889',
    ComplianceAssetId: null,
    Title: null,
    ID: 6,
    Created: '2018-02-05T09:42:36',
    AuthorId: 1,
    Modified: '2018-02-05T09:44:03',
    EditorId: 1,
    OData__CopySource: null,
    CheckoutUserId: null,
    OData__UIVersionString: '3.0',
    GUID: '9a94cb88-019b-4a66-abd6-be7f5337f659'
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/article.aspx')/ListItemAllFields`) > -1 ||
        (opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sitepages/article.aspx')/ListItemAllFields`) > -1) &&
        JSON.stringify(opts.data) === JSON.stringify({
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/article.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return;
      }

      throw 'Invalid request';
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      spo.getFileAsListItemByUrl,
      spo.systemUpdateListItem
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PAGE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates page layout to Article', async () => {
    sinon.stub(spo, 'systemUpdateListItem').resolves();
    sinon.stub(spo, 'getFileAsListItemByUrl').resolves(fileResponse);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: false, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' }) });
    assert(loggerLogSpy.notCalled);
  });

  it('updates page layout to Article (debug)', async () => {
    sinon.stub(spo, 'systemUpdateListItem').resolves();
    sinon.stub(spo, 'getFileAsListItemByUrl').resolves(fileResponse);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' }) });
  });

  it('updates page layout to Article on root of tenant (debug)', async () => {
    sinon.stub(spo, 'systemUpdateListItem').resolves();
    sinon.stub(spo, 'getFileAsListItemByUrl').resolves(fileResponse);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' }) });
  });

  it('automatically appends the .aspx extension', async () => {
    sinon.stub(spo, 'systemUpdateListItem').resolves();
    sinon.stub(spo, 'getFileAsListItemByUrl').resolves(fileResponse);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: false, name: 'article', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' }) });
    assert(loggerLogSpy.notCalled);
  });

  it('updates page layout to Home', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(spo, 'systemUpdateListItem').resolves();
    sinon.stub(spo, 'getFileAsListItemByUrl').resolves(fileResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage") {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/team-a/_api/SitePages/Pages(1)/SavePageAsDraft') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home' }) });
    assert(loggerLogSpy.notCalled);
  });

  it('promotes the page as NewsPage', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(spo, 'systemUpdateListItem').resolves();
    sinon.stub(spo, 'getFileAsListItemByUrl').resolves(fileResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        opts.data.PromotedState === 2 &&
        opts.data.FirstPublishedDate) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', promoteAs: 'NewsPage' }) });
    assert(loggerLogSpy.notCalled);
  });

  it('promotes the page as Template', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(spo, 'systemUpdateListItem').resolves();
    sinon.stub(spo, 'getFileAsListItemByUrl').resolves(fileResponse);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team-a/_api/SitePages/Pages(6)/SavePageAsTemplate`) {
        return { Id: 2, BannerImageUrl: 'url', CanvasContent1: 'content1', LayoutWebpartsContent: 'content' };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/team-a/_api/SitePages/Pages(2)/SavePageAsDraft`) {
        return;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/team-a/_api/SitePages/Pages(2)/SavePage`) {
        return;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) {
        return {
          Id: 6
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', description: "template", promoteAs: 'Template' }) });
    assert(loggerLogSpy.notCalled);
    assert.strictEqual(postStub.callCount, 4);
  });

  it('correctly marks the page as retired', async () => {
    sinonUtil.restore([request.post]);

    const systemUpdateItemStub = sinon.stub(spo, 'systemUpdateListItem').callsFake(async (requestUrl: string, listItemId: number) => {
      if (requestUrl === `https://contoso.sharepoint.com/sites/team-a/_api/web/GetList('${formatting.encodeQueryParameter('/sites/team-a/sitepages')}')` && listItemId === 6) {
        return {} as ListItemInstance;
      }

      throw 'Invalid systemUpdateListItem call.';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team-a/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1=''&@a2=1`) {
        return;
      }

      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage") {
        return {
          Title: "article",
          Id: 6,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', isRetired: true, publish: true }) });
    assert.deepStrictEqual(systemUpdateItemStub.lastCall.args[4], { _SPIsRetired: true });
    assert(postStub.calledTwice);
  });

  it('updates page layout to Home and promotes it as HomePage (debug)', async () => {
    sinonUtil.restore([request.post]);
    sinon.stub(spo, 'systemUpdateListItem').resolves();
    sinon.stub(spo, 'getFileAsListItemByUrl').resolves(fileResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          PageLayoutType: 'Home'
        })) {
        return;
      }

      if ((opts.url as string).indexOf('_api/web/rootfolder') > -1 &&
        opts.data.WelcomePage === 'SitePages/page.aspx') {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home', promoteAs: 'HomePage', description: "Page Description" }) });
  });

  it('enables comments on the page', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(false)`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: true }) });
  });

  it('disables comments on the page (debug)', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage") {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields" ||
        opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(true)" ||
        opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/SitePages/Pages(1)/SavePageAsDraft") {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: false }) });
    assert(loggerLogSpy.notCalled);
  });

  it('demotes news page to a regular page', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(spo, 'systemUpdateListItem').resolves();
    sinon.stub(spo, 'getFileAsListItemByUrl').resolves(fileResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) {
        return {};
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', demoteFrom: 'NewsPage' }) });
  });

  it('updates page title', async () => {
    sinonUtil.restore([request.post]);

    const newPageTitle = "updated title";

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(false)`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: newPageTitle }) });
  });

  it('updates page content', async () => {
    sinonUtil.restore([request.post]);

    const newContent = [{
      "controlType": 4,
      "position": {
        "layoutIndex": 1,
        "zoneIndex": 1,
        "sectionIndex": 1,
        "sectionFactor": 12,
        "controlIndex": 1
      },
      "addedFromPersistedData": true,
      "innerHTML": "<p>Text content</p>"
    }];

    const initialPage = {
      Title: "article",
      Id: 1,
      TopicHeader: "TopicHeader",
      AuthorByline: "AuthorByline",
      Description: "Description",
      BannerImageUrl: {
        Description: '/_layouts/15/images/sitepagethumbnail.png',
        Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
      },
      CanvasContent1: "{}",
      LayoutWebpartsContent: "{}"
    };

    let data: string = '';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).includes(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`)) {
        return initialPage;
      }
      if ((opts.url as string).includes(`/_api/SitePages/Pages(1)/SavePage`) ||
        (opts.url as string).includes(`/_api/SitePages/Pages(1)/SavePageAsDraft`)) {
        data = opts.data.CanvasContent1;
        return {};
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', content: JSON.stringify(newContent) }) });
    assert.deepStrictEqual(JSON.parse(data), newContent);
  });

  it('publishes page', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1=''&@a2=1`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true }) });
    assert(loggerLogSpy.notCalled);
  });

  it('publishes page with a message (debug)', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1='Initial%20version'&@a2=1`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Initial version' }) });
  });

  it('escapes special characters in user input', async () => {
    sinonUtil.restore([request.post]);
    const comment = `Don't tell`;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/team-a/_api/web/GetFileByServerRelativePath(DecodedUrl='/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${formatting.encodeQueryParameter(comment)}'&@a2=1`) {
        return;
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return {
          Title: "article",
          Id: 1,
          TopicHeader: "TopicHeader",
          AuthorByline: "AuthorByline",
          Description: "Description",
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          },
          CanvasContent1: "{}",
          LayoutWebpartsContent: "{}"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: comment }) });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles OData error when creating modern page', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake(async () => {
      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' }) }),
      new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not an absolute URL', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'foo' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'http://foo' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `'http://foo' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation when name and webURL specified and webUrl is a valid SharePoint URL but no option is set to update', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when name has no extension', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: true });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if layout type is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if layout type is Home', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Home' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if layout type is Article', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if layout type is SingleWebPartAppPage', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'SingleWebPartAppPage' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if layout type is RepostPage', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'RepostPage' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if layout type is HeaderlessSearchResults', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'HeaderlessSearchResults' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if layout type is Spaces', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Spaces' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if layout type is Topic', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Topic' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if demote type is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', demoteFrom: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if promote type is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if promote type is HomePage', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Home' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if promote type is NewsPage', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if promote type is Template', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'Template' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if promote type is HomePage but layout type is not Home', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Article' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if promote type is NewsPage but layout type is not Article', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage', layoutType: 'Home' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if commentsEnabled is true', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if commentsEnabled is false', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: false });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if content is not valid JSON', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', content: "foo" });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'Specified content is not a valid JSON string.');
  });

  it('passes validation if content is valid JSON', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', content: "[]" });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});