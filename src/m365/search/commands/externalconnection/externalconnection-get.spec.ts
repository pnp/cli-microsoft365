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

  const externalConnection: any =
  {
    "id": "contosohr",
    "name": "Contoso HR",
    "description": "Connection to index Contoso HR system",
    "state": "draft",
    "configuration": {
      "authorizedApps": [
        "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
      ],
      "authorizedAppIds": [
        "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
      ]
    }
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
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
  
  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['id', 'name']]);
  });
  
  it('correctly handles error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: {
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get external connection information for the Microsoft Search by id (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/contosohr`) {
        return Promise.resolve(externalConnection);
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        id: 'contosohr'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
        assert.strictEqual(call.args[0].id, 'contosohr');
        assert.strictEqual(call.args[0].name, 'Contoso HR');
        assert.strictEqual(call.args[0].description, 'Connection to index Contoso HR system');
        assert.strictEqual(call.args[0].state, 'draft');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get external connection information for the Microsoft Search by name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq '`) > -1) {
        return Promise.resolve({
          "value": [
            externalConnection
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso HR'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
        assert.strictEqual(call.args[0].id, 'contosohr');
        assert.strictEqual(call.args[0].name, 'Contoso HR');
        assert.strictEqual(call.args[0].description, 'Connection to index Contoso HR system');
        assert.strictEqual(call.args[0].state, 'draft');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails retrieving external connection not found by name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/external/connections?$filter=name eq '`) > -1) {
        return Promise.resolve({
          "value": []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso HR'
      }
    }, (err?: any) => {
      try {
        assert.deepStrictEqual(err, new CommandError(`External connection with name 'Contoso HR' not found`));
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
});