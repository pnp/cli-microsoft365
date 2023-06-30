import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError } from '../../../Command.js';
import request from '../../../request.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command from './app-get.js';
import { aadApp } from '../../../utils/aadApp.js';

describe(commands.GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({
      "apps": [
        {
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "name": "CLI app1"
        }
      ]
    }));
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      aadApp.getAppById
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles error when the app specified with the appId not found', async () => {
    const error = `No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`;
    sinon.stub(aadApp, 'getAppById').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    }), new CommandError(`No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`));
  });

  it(`gets an Azure AD app registration by its app (client) ID.`, async () => {
    const appResponse = {
      value: [
        {
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        }
      ]
    };

    sinon.stub(aadApp, 'getAppById').resolves(appResponse.value[0]);

    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    });
    const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
    assert.strictEqual(call.args[0].id, '340a4aa3-1af6-43ac-87d8-189819003952');
    assert.strictEqual(call.args[0].appId, '9b1b1e42-794b-4c71-93ac-5ed92488b67f');
    assert.strictEqual(call.args[0].displayName, 'My App');
  });
});