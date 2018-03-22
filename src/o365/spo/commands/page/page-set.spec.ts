import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./page-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.PAGE_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigestForSite').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      (command as any).getRequestDigestForSite
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.PAGE_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.PAGE_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates page layout to Article', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates page layout to Article (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('automatically appends the .aspx extension', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          PageLayoutType: 'Article',
          PromotedState: 0,
          BannerImageUrl: {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png`
          }
        })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, name: 'page', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates page layout to Home', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          PageLayoutType: 'Home'
        })) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('promotes the page as NewsPage', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        opts.body.PromotedState === 2 &&
        opts.body.FirstPublishedDate) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', promoteAs: 'NewsPage' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates page layout to Home and promotes it as HomePage (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          PageLayoutType: 'Home'
        })) {
        return Promise.resolve();
      }

      if (opts.url.indexOf('_api/web/rootfolder') > -1 &&
        opts.body.WelcomePage === 'SitePages/page.aspx') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Home', promoteAs: 'HomePage' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('enables comments on the page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(false)`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: 'true' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables comments on the page (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields/SetCommentsDisabled(true)`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', commentsEnabled: 'false' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('publishes page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/Publish('')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('publishes page with a message (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/Publish('Initial%20version')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Initial version' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes special characters in user input', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/Publish('Don%39t%20tell')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', publish: true, publishMessage: 'Don\'t tell' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when creating modern page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com/sites/team-a', layoutType: 'Article' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying name', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page layout', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--layoutType') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page promote option', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--promoteAs') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if comments should be enabled', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--commentsEnabled') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying if page should be published', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publish') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page publish message', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--publishMessage') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if name not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not an absolute URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://foo.com' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' } });
    assert.equal(actual, true);
  });

  it('passes validation when name has no extension', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page', webUrl: 'https://contoso.sharepoint.com' } });
    assert.equal(actual, true);
  });

  it('fails validation if layout type is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if layout type is Home', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Home' } });
    assert.equal(actual, true);
  });

  it('passes validation if layout type is Article', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', layoutType: 'Article' } });
    assert.equal(actual, true);
  });

  it('fails validation if promote type is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if promote type is HomePage', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Home' } });
    assert.equal(actual, true);
  });

  it('passes validation if promote type is NewsPage', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage' } });
    assert.equal(actual, true);
  });

  it('fails validation if promote type is HomePage but layout type is not Home', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'HomePage', layoutType: 'Article' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if promote type is NewsPage but layout type is not Article', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', promoteAs: 'NewsPage', layoutType: 'Home' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if commentsEnabled is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if commentsEnabled is true', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: 'true' } });
    assert.equal(actual, true);
  });

  it('passes validation if commentsEnabled is false', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', commentsEnabled: 'false' } });
    assert.equal(actual, true);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.PAGE_SET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com', name: 'page.aspx' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});