import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./hubsite-data-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.HUBSITE_DATA_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X', forceRefresh: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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

  it('gets information about the specified hub site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/HubSiteData`) > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(`https://contoso.sharepoint.com/sites/Project-X is not connected to a hub site and is not a hub site itself`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when hub site not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
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

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/Project-X' } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
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

  it('supports specifying forceRefresh', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--forceRefresh') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webUrl: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passed validation if webUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9b142c22-037f-4a7f-9017-e9d8c0e34b99', webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });
});