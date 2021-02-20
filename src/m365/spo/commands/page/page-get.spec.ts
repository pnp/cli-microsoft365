import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import { controlsMock, pageListItemMock, sectionMock, classicPage } from './page-get.mock';
const command: Command = require('./page-get');

describe(commands.PAGE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['commentsDisabled', 'numSections', 'numControls', 'title', 'layoutType']);
  });

  it('gets information about a modern page including all returned properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(pageListItemMock);
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return Promise.resolve(controlsMock);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx', output: 'json'}} as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].numControls, sectionMock.numControls);
        assert.strictEqual(loggerLogSpy.lastCall.args[0].numSections, sectionMock.numSections);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about a modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(pageListItemMock);
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return Promise.resolve(controlsMock);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx'}} as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
          ...pageListItemMock,
          canvasContentJson: controlsMock.CanvasContent1,
          ...sectionMock
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about a modern page on root of tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(pageListItemMock);
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return Promise.resolve(controlsMock);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx', output: 'json' }} as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
          ...pageListItemMock,
          canvasContentJson: controlsMock.CanvasContent1,
          ...sectionMock
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about a modern page when the specified page name doesn\'t contain extension', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(pageListItemMock);
      }

      if ((opts.url as string).indexOf(`/_api/SitePages/Pages(83)`) > -1) {
        return Promise.resolve(controlsMock);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home', output: 'json' }} as any, () => {
      try {
        assert(loggerLogSpy.calledWith({
          ...pageListItemMock,
          canvasContentJson: controlsMock.CanvasContent1,
          ...sectionMock
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('check if section and control HTML parsing gets skipped for metadata only mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(pageListItemMock);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home', metadataOnly: true, output: 'json' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(pageListItemMock));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows error when the specified page is a classic page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/SitePages/home.aspx')`) > -1) {
        return Promise.resolve(classicPage);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Page home.aspx is not a modern page.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles page not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({ error: {
        "odata.error": {
          "code": "-2130575338, Microsoft.SharePoint.SPException",
          "message": {
            "lang": "en-US",
            "value": "The file /sites/team-a/SitePages/home1.aspx does not exist."
          }
        }
      } });
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The file /sites/team-a/SitePages/home1.aspx does not exist.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving pages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', name: 'home.aspx' } } as any, (err?: any) => {
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

  it('supports specifying metadataOnly flag', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--metadataOnly') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo', name: 'home.aspx' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL and name is specified', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', name: 'home.aspx' } });
    assert.strictEqual(actual, true);
  });
});
