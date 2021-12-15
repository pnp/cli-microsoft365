import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./externalconnection-add');

describe(commands.EXTERNALCONNECTION_ADD, () => {
  let log: string[];
  let logger: Logger;

  const externalConnectionAddResponse: any = {
    configuration: {
      authorizedAppIds: []
    },
    description: 'Test connection that will not do anything',
    id: 'TestConnectionForCLI',
    name: 'Test Connection for CLI'
  };

  const externalConnectionAddResponseWithAppIDs: any = {
    configuration: {
      'authorizedAppIds': [
        '00000000-0000-0000-0000-000000000000',
        '00000000-0000-0000-0000-000000000001',
        '00000000-0000-0000-0000-000000000002'
      ]
    },
    description: 'Test connection that will not do anything',
    id: 'TestConnectionForCLI',
    name: 'Test Connection for CLI'
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    (command as any).items = [];
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
    assert.strictEqual(command.name.startsWith(commands.EXTERNALCONNECTION_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds an external connection', (done: any) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
        return Promise.resolve(externalConnectionAddResponse);
      }
      return Promise.reject('Invalid request');
    });
    const options: any = {
      debug: false,
      id: 'TestConnectionForCLI',
      name: 'Test Connection for CLI',
      description: 'Test connection that will not do anything'
    };
    command.action(logger, { options: options } as any, () => {
      try {
        assert.deepStrictEqual(postStub.getCall(0).args[0].data, externalConnectionAddResponse);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds an external connection with authorized app id', (done: any) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
        return Promise.resolve(externalConnectionAddResponse);
      }
      return Promise.reject('Invalid request');
    });
    const options: any = {
      debug: false,
      id: 'TestConnectionForCLI',
      name: 'Test Connection for CLI',
      description: 'Test connection that will not do anything',
      authorizedAppIds: '00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002'
    };
    command.action(logger, { options: options } as any, () => {
      try {
        assert.deepStrictEqual(postStub.getCall(0).args[0].data, externalConnectionAddResponseWithAppIDs);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds an external connection with authorised app IDs', (done: any) => {
    const postStub = sinon.stub(request, 'post').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
        return Promise.resolve(externalConnectionAddResponseWithAppIDs);
      }
      return Promise.reject('Invalid request');
    });
    const options: any = {
      debug: false,
      id: 'TestConnectionForCLI',
      name: 'Test Connection for CLI',
      description: 'Test connection that will not do anything',
      authorizedAppIds: '00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000001,00000000-0000-0000-0000-000000000002'
    };
    command.action(logger, { options: options } as any, () => {
      try {
        assert.deepStrictEqual(postStub.getCall(0).args[0].data, externalConnectionAddResponseWithAppIDs);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "Error",
          "message": "An error has occurred",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if id is less than 3 characters', (done) => {
    const actual = command.validate({
      options: {
        id: 'T',
        name: 'Test Connection for CLI'
      }
    });
    assert.notStrictEqual(actual, false);
    done();
  });

  it('fails validation if id is more than 32 characters', (done) => {
    const actual = command.validate({
      options: {
        id: 'TestConnectionForCLIXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX',
        name: 'Test Connection for CLI'
      }
    });
    assert.notStrictEqual(actual, false);
    done();
  });

  it('fails validation if id is not alphanumeric', (done) => {
    const actual = command.validate({
      options: {
        id: 'Test_Connection!',
        name: 'Test Connection for CLI'
      }
    });
    assert.notStrictEqual(actual, false);
    done();
  });

  it('fails validation if id starts with Microsoft', () => {
    const actual = command.validate({
      options: {
        id: 'MicrosoftTestConnectionForCLI',
        name: 'Test Connection for CLI'
      }
    });
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if id is SharePoint', () => {
    const actual = command.validate({
      options: {
        id: 'SharePoint',
        name: 'Test Connection for CLI'
      }
    });
    assert.notStrictEqual(actual, false);
  });

  it('passes validation for a valid id', () => {
    const actual = command.validate({
      options: {
        id: 'myapp',
        name: 'Test Connection for CLI'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports specifying id', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
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

  it('supports specifying description', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
