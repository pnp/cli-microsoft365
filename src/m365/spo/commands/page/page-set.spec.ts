import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./page-set');

describe(commands.PAGE_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');

    sinon.stub(request, 'post').callsFake((opts) => {
      if (((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/article.aspx')/ListItemAllFields`) > -1 ||
        (opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sitepages/article.aspx')/ListItemAllFields`) > -1) &&
        JSON.stringify(opts.data) === JSON.stringify({
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/article.aspx')/checkoutpage`) > -1) {
        return Promise.resolve({
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
        });
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates page layout to Article', async () => {
    await command.action(logger, { options: { debug: false, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates page layout to Article (debug)', async () => {
    await command.action(logger, { options: { debug: true, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } });
  });

  it('updates page layout to Article on root of tenant (debug)', async () => {
    await command.action(logger, { options: { debug: true, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' } });
  });

  it('automatically appends the .aspx extension', async () => {
    await command.action(logger, { options: { debug: false, name: 'article', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates page layout to Home', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage") {
        return Promise.resolve({
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
        });
      }

      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields") {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home' } }));
    assert(loggerLogSpy.notCalled);
  });

  it('promotes the page as NewsPage', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        opts.data.PromotedState === 2 &&
        opts.data.FirstPublishedDate) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', promoteAs: 'NewsPage' } });
    assert(loggerLogSpy.notCalled);
  });

  it('promotes the page as Template', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        !opts.data) {
        return Promise.resolve({ Id: '1' });
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsTemplate`) > -1) {
        return Promise.resolve({ Id: '2', BannerImageUrl: 'url', CanvasContent1: 'content1', LayoutWebpartsContent: 'content' });
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(2)/SavePage`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', description: "template", promoteAs: 'Template' } } as any);
    assert(loggerLogSpy.notCalled);
  });

  it('updates page layout to Home and promotes it as HomePage (debug)', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          PageLayoutType: 'Home'
        })) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf('_api/web/rootfolder') > -1 &&
        opts.data.WelcomePage === 'SitePages/page.aspx') {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return Promise.resolve({
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
        });
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home', promoteAs: 'HomePage', description: "Page Description" } });
  });

  it('enables comments on the page', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(false)`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return Promise.resolve({
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
        });
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: 'true' } });
  });

  it('disables comments on the page (debug)', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage") {
        return Promise.resolve({
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
        });
      }

      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields" ||
        opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(true)" ||
        opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/SitePages/Pages(1)/SavePageAsDraft") {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: 'false' } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates page title', async () => {
    sinonUtil.restore([request.post]);

    const newPageTitle = "updated title";

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(false)`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return Promise.resolve({
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
        });
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: newPageTitle } });
  });

  it('publishes page', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return Promise.resolve({
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
        });
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1=\'\'&@a2=1`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('publishes page with a message (debug)', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return Promise.resolve({
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
        });
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePage`) > -1) {
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1='Initial%20version'&@a2=1`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Initial version' } });
  });

  it('escapes special characters in user input', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/team-a/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/CheckIn(comment=@a1,checkintype=@a2)?@a1='Don%39t%20tell'&@a2=1") {
        return Promise.resolve();
      }
      
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/checkoutpage`) > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Don\'t tell' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles OData error when creating modern page', async () => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    await assert.rejects(command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying name', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page layout', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--layoutType') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page promote option', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--promoteAs') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if comments should be enabled', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--commentsEnabled') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if page should be published', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publish') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page publish message', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publishMessage') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'http://foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name has no extension', async () => {
    const actual = await command.validate({ options: { name: 'page', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if layout type is invalid', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if layout type is Home', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Home' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is Article', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is SingleWebPartAppPage', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'SingleWebPartAppPage' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is RepostPage', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'RepostPage' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is HeaderlessSearchResults', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'HeaderlessSearchResults' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is Spaces', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Spaces' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is Topic', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Topic' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if promote type is invalid', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if promote type is HomePage', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Home' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if promote type is NewsPage', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if promote type is Template', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'Template' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if promote type is HomePage but layout type is not Home', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Article' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if promote type is NewsPage but layout type is not Article', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage', layoutType: 'Home' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if commentsEnabled is invalid', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if commentsEnabled is true', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if commentsEnabled is false', async () => {
    const actual = await command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
