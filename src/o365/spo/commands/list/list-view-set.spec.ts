import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./list-view-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.LIST_VIEW_SET, () => {
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
      request.patch
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      request.patch
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LIST_VIEW_SET), true);
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
        assert.equal(telemetry.name, commands.LIST_VIEW_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the Title of the list view specified using its name', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('List%201')/views/getByTitle('All%20items')`) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.body) === JSON.stringify({ Title: 'All events' })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewTitle: 'All items', Title: 'All events' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates the Title and CustomFormatter of the list view specified using its ID', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'330f29c5-5c4c-465f-9f4b-7903020ae1cf')/views/getById('330f29c5-5c4c-465f-9f4b-7903020ae1ce')`) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.body) === JSON.stringify({ Title: 'All events', CustomFormatter: 'abc' })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com', listId: '330f29c5-5c4c-465f-9f4b-7903020ae1cf', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', Title: 'All events', CustomFormatter: 'abc' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified list doesn\'t exist', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-1, System.ArgumentException",
            "message": {
              "lang": "en-US",
              "value": "List 'List' does not exist at site with URL 'https://contoso.sharepoint.com'."
            }
          }
        }
      })
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List', viewTitle: 'All items' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("List 'List' does not exist at site with URL 'https://contoso.sharepoint.com'.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified view doesn\'t exist', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      return Promise.reject({
        error: {
          "odata.error": {
            "code": "-2147024809, System.ArgumentException",
            "message": {
              "lang": "en-US",
              "value": "The specified view is invalid."
            }
          }
        }
      })
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'List', viewTitle: 'All items' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("The specified view is invalid.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('allows unknown options', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.equal(allowUnknownOptions, true);
  });

  it('fails validation if webUrl is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { listTitle: 'List 1', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid', listTitle: 'List 1', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if neither listId nor listTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if both listId and listTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', listId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if listId is not a GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'invalid', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if neither viewId nor viewTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if both viewId and viewTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewTitle: 'All items' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if viewId is not a GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewId: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when viewTitle and listTitle specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewTitle: 'All items' } });
    assert.equal(actual, true);
  });

  it('passes validation when viewId and listId specified and valid GUIDs', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', viewId: '330f29c5-5c4c-465f-9f4b-7903020ae1cf' } });
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
    assert(find.calledWith(commands.LIST_VIEW_SET));
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
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});