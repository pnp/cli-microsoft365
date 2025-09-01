import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './managementapp-list.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.MANAGEMENTAPP_LIST, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
    auth.connection.active = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MANAGEMENTAPP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('successfully retrieves management application', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return {
          "value": [{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }]
        };
      }
      throw 'Invalid request';
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return {
          "value": [{ "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548e" }, { "applicationId": "31359c7f-bd7e-475c-86db-fdb8c937548f" }]
        };
      }
      throw 'Invalid request';
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01") {
        return {
          "value": [{}]
        };
      }
      throw 'Invalid request';
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
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
