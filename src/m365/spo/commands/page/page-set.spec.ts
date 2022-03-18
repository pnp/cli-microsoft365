import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./page-set');

describe(commands.PAGE_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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

  it('updates page layout to Article', (done) => {
    command.action(logger, { options: { debug: false, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates page layout to Article (debug)', (done) => {
    command.action(logger, { options: { debug: true, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates page layout to Article on root of tenant (debug)', (done) => {
    command.action(logger, { options: { debug: true, name: 'article.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('automatically appends the .aspx extension', (done) => {
    command.action(logger, { options: { debug: false, name: 'article', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates page layout to Home', (done) => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          PageLayoutType: 'Home'
        })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('promotes the page as NewsPage', (done) => {
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

    command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', promoteAs: 'NewsPage' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('promotes the page as Template', (done) => {
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

    command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', description: "template", promoteAs: 'Template' } } as any, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates page layout to Home and promotes it as HomePage (debug)', (done) => {
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

    command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home', promoteAs: 'HomePage', description: "Page Description" } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('enables comments on the page', (done) => {
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

    command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: 'true' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables comments on the page (debug)', (done) => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(true)`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: 'false' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates page title', (done) => {
    sinonUtil.restore([request.post]);

    const newPageTitle = "updated title";
    let responseData: any = {};

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
        responseData = opts.data;
        return Promise.resolve();
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(1)/SavePageAsDraft`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', title: newPageTitle } }, () => {
      try {
        assert.strictEqual(responseData.Title, newPageTitle);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('publishes page', (done) => {
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

    command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('publishes page with a message (debug)', (done) => {
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

    command.action(logger, { options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Initial version' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes special characters in user input', (done) => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/Publish('Don%39t%20tell')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Don\'t tell' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when creating modern page', (done) => {
    sinonUtil.restore([request.post]);

    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying name', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page layout', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--layoutType') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page promote option', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--promoteAs') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if comments should be enabled', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--commentsEnabled') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if page should be published', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publish') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page publish message', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publishMessage') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'http://foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name has no extension', () => {
    const actual = command.validate({ options: { name: 'page', webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if layout type is invalid', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if layout type is Home', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Home' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is Article', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is SingleWebPartAppPage', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'SingleWebPartAppPage' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is RepostPage', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'RepostPage' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is HeaderlessSearchResults', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'HeaderlessSearchResults' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is Spaces', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Spaces' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if layout type is Topic', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Topic' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if promote type is invalid', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if promote type is HomePage', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Home' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if promote type is NewsPage', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if promote type is Template', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'Template' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if promote type is HomePage but layout type is not Home', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Article' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if promote type is NewsPage but layout type is not Article', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage', layoutType: 'Home' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if commentsEnabled is invalid', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if commentsEnabled is true', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: 'true' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if commentsEnabled is false', () => {
    const actual = command.validate({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: 'false' } });
    assert.strictEqual(actual, true);
  });
});
