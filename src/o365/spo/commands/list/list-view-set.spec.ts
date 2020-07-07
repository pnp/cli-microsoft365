import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-view-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_VIEW_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
      request.patch
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
    assert.equal(command.name.startsWith(commands.LIST_VIEW_SET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('updates the Title of the list view specified using its name', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('List%201')/views/getByTitle('All%20items')`) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.body) === JSON.stringify({ Title: 'All events' })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

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

  it('ignores global options when creating request body', (done) => {
    const patchRequest: sinon.SinonStub = sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('List%201')/views/getByTitle('All%20items')`) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.body) === JSON.stringify({ Title: 'All events' })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, verbose: false, output: "text", webUrl: 'https://contoso.sharepoint.com', listTitle: 'List 1', viewTitle: 'All items', Title: 'All events' } }, () => {
      try {
        assert.deepEqual(patchRequest.lastCall.args[0].body, { Title: 'All events' });
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
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0 &&
          opts.headers['X-RequestDigest'] &&
          JSON.stringify(opts.body) === JSON.stringify({ Title: 'All events', CustomFormatter: 'abc' })) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

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
});