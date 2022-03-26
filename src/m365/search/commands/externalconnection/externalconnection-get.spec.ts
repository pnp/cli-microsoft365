import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./externalconnection-get');

describe(commands.EXTERNALCONNECTION_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const externalConnectionGetResponse: any = {
    configuration: {
      authorizedAppIds: []
    },
    description: 'Test connection that will not do anything',
    id: 'TestConnectionForCLI',
    name: 'Test Connection for CLI'
  };

  const externalConnectionGetResponseWithAppIDs: any = {
    configuration: {
      'authorizedAppIds': [
        '00000000-0000-0000-0000-000000000000',
        '00000000-0000-0000-0000-000000000001',
        '00000000-0000-0000-0000-000000000002'
      ]
    },
    description: 'Test connection that will not do anything',
    id: 'TestConnectionForCLIWithAppIDs',
    name: 'Test Connection for CLI with App IDs'
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
    
    loggerLogSpy = sinon.spy(logger, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.EXTERNALCONNECTION_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets an external connection', (done: any) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/TestConnectionForCLI`) {
        return Promise.resolve(externalConnectionGetResponse);
      }
      return Promise.reject('Invalid request');
    });
    const options: any = {
      debug: false,
      id: 'TestConnectionForCLI',
      name: 'Test Connection for CLI',
      description: 'Test connection that will not do anything'
    };

    command.action(logger, { options: options }, () => {
      try {
        assert(loggerLogSpy.calledWith(externalConnectionGetResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets an external connection with authorized app id', (done: any) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/TestConnectionForCLIWithAppIDs`) {
        return Promise.resolve(externalConnectionGetResponseWithAppIDs);
      }
      return Promise.reject('Invalid request');
    });
    const options: any = {
      debug: false,
      id: 'TestConnectionForCLIWithAppIDs',
      name: 'Test Connection for CLI With App IDs',
      description: 'Test connection that will not do anything'
    };

    command.action(logger, { options: options }, () => {
      try {
        assert(loggerLogSpy.calledWith(externalConnectionGetResponseWithAppIDs));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get external connection when team does not exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      const matchingItemNumber = (opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq 'sjalfj'`);
      
      if (matchingItemNumber > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('The specified external connection does not exist');
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'Test App'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified external connection does not exist`)));
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
        id: 'T'
      }
    });
    assert.notStrictEqual(actual, false);
    done();
  });

  it('fails validation if id is more than 32 characters', (done) => {
    const actual = command.validate({
      options: {
        id: 'TestConnectionForCLIXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
      }
    });
    assert.notStrictEqual(actual, false);
    done();
  });

  it('fails validation if id is not alphanumeric', (done) => {
    const actual = command.validate({
      options: {
        id: 'Test_Connection!'
      }
    });
    assert.notStrictEqual(actual, false);
    done();
  });

  it('fails validation if id starts with Microsoft', () => {
    const actual = command.validate({
      options: {
        id: 'MicrosoftTestConnectionForCLI'
      }
    });
    assert.notStrictEqual(actual, false);
  });

  it('fails validation if id is SharePoint', () => {
    const actual = command.validate({
      options: {
        id: 'SharePoint'
      }
    });
    assert.notStrictEqual(actual, false);
  });

  it('passes validation for a valid id', () => {
    const actual = command.validate({
      options: {
        id: 'myapp'
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
});
