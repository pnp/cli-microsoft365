import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./externalconnection-list');

describe(commands.EXTERNALCONNECTION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const externalConnectionListResponse: any = {
    configuration: {
      authorizedAppIds: []
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
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

  
  it('correctly handles error', (done) => {
    logger.logRaw('testing the error');

    sinon.stub(request, 'get').callsFake(() => {
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

    command.action(logger, {
      options: {
        debug: false
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(externalConnectionListResponse));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('lists an external connection', (done) => {
    const postStub = sinon.stub(request, 'get').callsFake((opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections`) {
        return Promise.resolve(externalConnectionListResponse);
      }
      return Promise.reject('Invalid request');
    });
    
    const options: any = {
      debug: false,
      externalConnectionId: 'TestConnectionForCLI',
      externalConnectionName: 'Test Connection for CLI',
      externalConnectionDescription: 'Test connection that will not do anything'
    };
  
    command.action(logger, { options: options } as any, () => {
      try {
        assert.deepStrictEqual(postStub.getCall(0).args[0].data, externalConnectionListResponse);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });


});