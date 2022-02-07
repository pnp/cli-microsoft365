import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./managementapp-list');

describe(commands.MANAGEMENTAPP_LIST, () => {
  let log: string[];
  let logger: Logger;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.put
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
    assert.strictEqual(command.name.startsWith(commands.MANAGEMENTAPP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('successfully retrieves management application', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return Promise.resolve({
          "value": [{"applicationId":"31359c7f-bd7e-475c-86db-fdb8c937548e"}]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        verbose: true
      }
    }, () => {
      try {
        const actual = JSON.stringify(log[log.length - 1]);
        const expected = JSON.stringify([{"applicationId":"31359c7f-bd7e-475c-86db-fdb8c937548e"}]);

        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('successfully retrieves multiple management applications', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return Promise.resolve({
          "value": [{"applicationId":"31359c7f-bd7e-475c-86db-fdb8c937548e"},{"applicationId":"31359c7f-bd7e-475c-86db-fdb8c937548f"}]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        verbose: true
      }
    }, () => {
      try {
        const actual = JSON.stringify(log[log.length - 1]);
        const expected = JSON.stringify([{"applicationId":"31359c7f-bd7e-475c-86db-fdb8c937548e"}, {"applicationId":"31359c7f-bd7e-475c-86db-fdb8c937548f"}]);

        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('successfully handles no result found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return Promise.resolve({
          "value": [ {} ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        verbose: true
      }
    }, () => {
      try {
        const actual = JSON.stringify(log[log.length - 1]);
        const expected = JSON.stringify([ {} ]);
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
