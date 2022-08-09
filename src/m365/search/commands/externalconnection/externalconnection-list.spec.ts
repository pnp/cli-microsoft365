import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./externalconnection-list');

describe(commands.EXTERNALCONNECTION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const externalConnections: any = {
    value: [
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
      }
    ]
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
    assert.strictEqual(command.name.startsWith(commands.EXTERNALCONNECTION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'state']);
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

  it('retrieves list of external connections defined in the Microsoft Search', (done) => {
    sinon.stub(request, 'get').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
        return Promise.resolve(externalConnections);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true } } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(externalConnections.value));
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