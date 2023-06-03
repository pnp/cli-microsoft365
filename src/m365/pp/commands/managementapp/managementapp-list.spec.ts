import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./managementapp-list');

describe(commands.MANAGEMENTAPP_LIST, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.MANAGEMENTAPP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('successfully retrieves management application', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return Promise.resolve({
          "value": [{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }]
        });
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });
    const actual = JSON.stringify(log[log.length - 1]);
    const expected = JSON.stringify([{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }]);

    assert.strictEqual(actual, expected);
  });

  it('successfully retrieves multiple management applications', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return Promise.resolve({
          "value": [{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }, { "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548f" }]
        });
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });
    const actual = JSON.stringify(log[log.length - 1]);
    const expected = JSON.stringify([{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }, { "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548f" }]);

    assert.strictEqual(actual, expected);
  });

  it('successfully handles no result found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return Promise.resolve({
          "value": [{}]
        });
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });
    const actual = JSON.stringify(log[log.length - 1]);
    const expected = JSON.stringify([{}]);
    assert.strictEqual(actual, expected);
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
