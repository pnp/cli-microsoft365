import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./oauth2grant-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.OAUTH2GRANT_ADD, () => {
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
    assert.strictEqual(command.name.startsWith(commands.OAUTH2GRANT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds OAuth2 permission grant (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/oauth2PermissionGrants?api-version=1.6`) > -1) {
        if (opts.headers &&
          opts.headers['content-type'] &&
          opts.headers['content-type'].indexOf('application/json') === 0 &&
          opts.body.clientId === '6a7b1395-d313-4682-8ed4-65a6265a6320' &&
          opts.body.resourceId === '6a7b1395-d313-4682-8ed4-65a6265a6321' &&
          opts.body.scope === 'user_impersonation') {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6321', scope: 'user_impersonation' } }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds OAuth2 permission grant', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/myorganization/oauth2PermissionGrants?api-version=1.6`) > -1) {
        if (opts.headers &&
          opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['content-type'] &&
          opts.headers['content-type'].indexOf('application/json') === 0 &&
          opts.body.clientId === '6a7b1395-d313-4682-8ed4-65a6265a6320' &&
          opts.body.resourceId === '6a7b1395-d313-4682-8ed4-65a6265a6321' &&
          opts.body.scope === 'user_impersonation') {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6321', scope: 'user_impersonation' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
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

    cmdInstance.action({ options: { debug: false, clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6320', scope: 'user_impersonation' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the clientId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { clientId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the resourceId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when clientId, resourceId and scope are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { clientId: '6a7b1395-d313-4682-8ed4-65a6265a6320', resourceId: '6a7b1395-d313-4682-8ed4-65a6265a6320', scope: 'user_impersonation' } });
    assert.strictEqual(actual, true);
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

  it('supports specifying clientId', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--clientId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying resourceId', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--resourceId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying scope', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--scope') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});