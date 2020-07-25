import commands from '../commands';
import Command, { CommandOption, CommandError } from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
const command: Command = require('./flow-disable');
import * as assert from 'assert';
import request from '../../../request';
import Utils from '../../../Utils';

describe(commands.FLOW_DISABLE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.FLOW_DISABLE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('disables the specified flow (debug)', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments`) > -1) {

        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5' } }, () => {
      try {
        assert.notStrictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE'), -1);
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://management.azure.com/providers/Microsoft.ProcessSimple/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d/stop?api-version=2016-11-01');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables the specified flow as admin', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.ProcessSimple/environments`) > -1) {

        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d', environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', asAdmin: true } }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://management.azure.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/Default-d87a7535-dd31-4437-bfe1-95340acd55c5/flows/3989cb59-ce1a-4a5c-bb78-257c5c39381d/stop?api-version=2016-11-01');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no environment found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "EnvironmentAccessDenied",
          "message": "Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Access to the environment 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6' is denied.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles Flow not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "ConnectionAuthorizationFailed",
          "message": "The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '1c6ee23a-a835-44bc-a4f5-462b658efc12' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The caller with object id 'da8f7aea-cf43-497f-ad62-c2feae89a194' does not have permission for connection '1c6ee23a-a835-44bc-a4f5-462b658efc12' under Api 'shared_logicflows'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles Flow not found (as admin)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "code": "FlowNotFound",
          "message": "Could not find flow '1c6ee23a-a835-44bc-a4f5-462b658efc12'."
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c6', name: '1c6ee23a-a835-44bc-a4f5-462b658efc12', asAdmin: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Could not find flow '1c6ee23a-a835-44bc-a4f5-462b658efc12'.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false, environment: 'Default-d87a7535-dd31-4437-bfe1-95340acd55c5', name: '3989cb59-ce1a-4a5c-bb78-257c5c39381d' } }, (err?: any) => {
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

  it('supports specifying environment', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--environment') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});