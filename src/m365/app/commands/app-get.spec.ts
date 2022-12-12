import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { telemetry } from '../../../telemetry';
import auth from '../../../Auth';
import { Logger } from '../../../cli/Logger';
import Command, { CommandError } from '../../../Command';
import request from '../../../request';
import { pid } from '../../../utils/pid';
import { sinonUtil } from '../../../utils/sinonUtil';
import commands from '../commands';
const command: Command = require('./app-get');

describe(commands.GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      fs.existsSync,
      fs.readFileSync
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles error when the app specified with the appId not found', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${JSON.stringify(opts)}`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    }), new CommandError(`No Azure AD application registration with ID 9b1b1e42-794b-4c71-93ac-5ed92488b67f found`));
  });

  it('handles error when retrieving information about app through appId failed', async () => {
    sinon.stub(request, 'get').callsFake(_ => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f'
      }
    }), new CommandError(`An error has occurred`));
  });

  it(`gets an Azure AD app registration by its app (client) ID.`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });

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

  it(`shows underlying debug information in debug mode`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=appId eq '9b1b1e42-794b-4c71-93ac-5ed92488b67f'&$select=id`) {
        return Promise.resolve({
          value: [
            {
              "id": "340a4aa3-1af6-43ac-87d8-189819003952",
              "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
              "createdDateTime": "2019-10-29T17:46:55Z",
              "displayName": "My App",
              "description": null
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/myorganization/applications/') > -1) {
        return Promise.resolve({
          "id": "340a4aa3-1af6-43ac-87d8-189819003952",
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "createdDateTime": "2019-10-29T17:46:55Z",
          "displayName": "My App",
          "description": null
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        debug: true
      }
    });
    const call: sinon.SinonSpyCall = loggerLogToStderrSpy.firstCall;
    assert(call.args[0].includes('Executing command aad app get with options'));
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
