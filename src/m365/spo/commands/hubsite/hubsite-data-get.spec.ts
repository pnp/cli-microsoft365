import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./hubsite-data-get');

describe(commands.HUBSITE_DATA_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HUBSITE_DATA_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about the specified hub site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData(false)`) > -1) {
        return Promise.resolve({
          value: JSON.stringify({
            "themeKey": null,
            "name": "CommunicationSite",
            "url": "https://contoso.sharepoint.com/sites/Sales",
            "logoUrl": "http://contoso.com/__siteIcon__.jpg",
            "usesMetadataNavigation": false,
            "navigation": []
          })
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "themeKey": null,
          "name": "CommunicationSite",
          "url": "https://contoso.sharepoint.com/sites/Sales",
          "logoUrl": "http://contoso.com/__siteIcon__.jpg",
          "usesMetadataNavigation": false,
          "navigation": []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site with forced refresh', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData(true)`) > -1) {
        return Promise.resolve({
          value: JSON.stringify({
            "themeKey": null,
            "name": "CommunicationSite",
            "url": "https://contoso.sharepoint.com/sites/Sales",
            "logoUrl": "http://contoso.com/__siteIcon__.jpg",
            "usesMetadataNavigation": false,
            "navigation": []
          })
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X', forceRefresh: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "themeKey": null,
          "name": "CommunicationSite",
          "url": "https://contoso.sharepoint.com/sites/Sales",
          "logoUrl": "http://contoso.com/__siteIcon__.jpg",
          "usesMetadataNavigation": false,
          "navigation": []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified hub site (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData`) > -1) {
        return Promise.resolve({
          value: JSON.stringify({
            "themeKey": null,
            "name": "CommunicationSite",
            "url": "https://contoso.sharepoint.com/sites/Sales",
            "logoUrl": "http://contoso.com/__siteIcon__.jpg",
            "usesMetadataNavigation": false,
            "navigation": []
          })
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "themeKey": null,
          "name": "CommunicationSite",
          "url": "https://contoso.sharepoint.com/sites/Sales",
          "logoUrl": "http://contoso.com/__siteIcon__.jpg",
          "usesMetadataNavigation": false,
          "navigation": []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles empty response', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData`) > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when specified site is not connect to or is a hub site (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData`) > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(`https://contoso.sharepoint.com/sites/Project-X is not connected to a hub site and is not a hub site itself`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when hub site not found', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-1, Microsoft.SharePoint.Client.ResourceNotFoundException",
            "message": {
              "lang": "en-US",
              "value": "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
            }
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('supports specifying forceRefresh', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--forceRefresh') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passed validation if webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});