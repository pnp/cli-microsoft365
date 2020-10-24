import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./group-remove');

describe(commands.GROUP_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let trackEvent: any;
  let telemetry: any;
  let promptOptions: any;

  before(() => {
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
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
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      request.get,
      Cli.prompt
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
    assert.strictEqual(command.name.startsWith(commands.GROUP_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls telemetry', () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, confirm: true } }, () => {
      assert(trackEvent.called);
    });
  });

  it('logs correct telemetry event', () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, confirm: true } }, () => {
      assert.strictEqual(telemetry.name, commands.GROUP_REMOVE);
    });
  });

  it('deletes the group when id is passed', (done) => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true, confirm: true } }, () => {
      try {
        assert(requestPostSpy.called);
        done();
      } catch (e) {
        done(e);
      }
    })
  });

  it('deletes the group when name is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/mysite/_api/web/sitegroups/GetByName('Team Site Owners')?$select=Id`) {
        return Promise.resolve({
          Id: 7
        });
      }
      return Promise.reject('Invalid Request');
    });

    const requestPostSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', name: 'Team Site Owners', debug: true, confirm: true } }, () => {
      try {
        assert(requestPostSpy.called);
        done();
      } catch (e) {
        done(e);
      }
    })
  });

  it('aborts deleting the group when prompt is not continued', (done) => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true } }, () => {
      try {
        assert(requestPostSpy.notCalled);
        done();
      } catch (e) {
        done(e);
      }
    })
  });

  it('deletes the group when prompt is continued', (done) => {
    const requestPostSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true } }, () => {
      try {
        assert(requestPostSpy.called);
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles group remove reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/mysite/_api/web/sitegroups/RemoveById(7)') {
        return Promise.reject(err);
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, debug: true, confirm: true } } as any, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      } catch (e) {
        done(e);
      }
    })
  });

  it('prompts before removing group when confirmation argument not passed (id)', (done) => {
    command.action(logger, { options: { debug: false, id: 7, webUrl: 'https://contoso.sharepoint.com/mysite' } }, () => {
      let promptIssued = false;
      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing group when confirmation argument not passed (name)', (done) => {
    command.action(logger, { options: { debug: false, name: 'Team Site Owners', webUrl: 'https://contoso.sharepoint.com/mysite' } }, () => {
      let promptIssued = false;
      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = command.options();
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if both id and name options are not passed', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7 } });
    assert(actual);
  });

  it('fails validation if the id option is not a number', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 'Hi' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a number', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7 } });
    assert(actual);
  });

  it('fails validation if both id and name options are passed', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/mysite', id: 7, name: 'Team Site Members' } });
    assert.notStrictEqual(actual, true);
  });
});